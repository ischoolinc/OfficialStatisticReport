using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SHSchool.Data;
using Aspose.Cells;
using System.IO;

namespace StudentAgeStatistics
{
    public partial class PrintSet : BaseForm
    {
        public PrintSet()
        {
            InitializeComponent();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            //運用Aspose元件來新增活頁簿

            

            Aspose.Cells.Workbook ScoreWorkBook = new Aspose.Cells.Workbook();

            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();

            saveFileDialog.FileName = "高中職學校班級及學生概況.xlsx";
            saveFileDialog.AddExtension = true;
            saveFileDialog.DefaultExt = "xlsx";

            string filename = (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) ? saveFileDialog.FileName : "";

            ScoreWorkBook.Open(new MemoryStream(Properties.Resource.高中職學校班級及學生概況));

            //if (ReportKind1.Checked == true)
            //    ScoreWorkBook.Open(new MemoryStream(Properties.Resource.高中職學校班級及學生概況_職業科));
            //if (ReportKind2.Checked == true)
            //    ScoreWorkBook.Open(new System.IO.MemoryStream(Properties.Resource.高中職學校班級及學生概況_普通科));
            //if (ReportKind3.Checked == true)
            //    ScoreWorkBook.Open(new System.IO.MemoryStream(Properties.Resource.高中職學校班級及學生概況_綜合高中));


            Program.ErrorList.Clear();

            #region 列印統計表
            ScoreWorkBook.Worksheets[0].Name = txtTitle.Text;
            ClassStatistics ClassSat = new ClassStatistics();
            //ComboBoxItem item = cboBranch.Items[cboBranch.SelectedIndex] as ComboBoxItem;
            List<string> CourseIDs=new List<string>();
            for (int i = 0; i < lstCourseKind.Items.Count; i++)
                if (lstCourseKind.Items[i].Checked == true)
                    CourseIDs.Add(lstCourseKind.Items[i].SubItems[1].Text);

            ClassSat.StartStatistics(CourseIDs);
            ScoreWorkBook.Worksheets[0].Cells[2, 0].PutValue("表-8 高級中等學校班級及學生概況(二)－" + txtTitle.Text);
            ScoreWorkBook.Worksheets[0].Cells[3, 8].PutValue(K12.Data.School.DefaultSchoolYear);
            ScoreWorkBook.Worksheets[0].Cells[4, 0].PutValue(K12.Data.School.Code);
            ScoreWorkBook.Worksheets[0].Cells[4, 3].PutValue(txtTitle.Text);
            ScoreWorkBook.Worksheets[0].Cells[0, 12].PutValue(K12.Data.School.ChineseName+"(教務處)");
            ////一年級班級數總計
            //ScoreWorkBook.Worksheets[0].Cells[10, 5].PutValue(ClassSat.Level1Count24);
            //ScoreWorkBook.Worksheets[0].Cells[10, 7].PutValue(ClassSat.Level1Count34);
            //ScoreWorkBook.Worksheets[0].Cells[10, 9].PutValue(ClassSat.Level1Count44);
            //ScoreWorkBook.Worksheets[0].Cells[10, 11].PutValue(ClassSat.Level1Count54);
            //ScoreWorkBook.Worksheets[0].Cells[10, 13].PutValue(ClassSat.Level1Count55);

            ////二年級班級數總計
            //ScoreWorkBook.Worksheets[0].Cells[11, 5].PutValue(ClassSat.Level2Count24);
            //ScoreWorkBook.Worksheets[0].Cells[11, 7].PutValue(ClassSat.Level2Count34);
            //ScoreWorkBook.Worksheets[0].Cells[11, 9].PutValue(ClassSat.Level2Count44);
            //ScoreWorkBook.Worksheets[0].Cells[11, 11].PutValue(ClassSat.Level2Count54);
            //ScoreWorkBook.Worksheets[0].Cells[11, 13].PutValue(ClassSat.Level2Count55);

            ////三年級班級數總計
            //ScoreWorkBook.Worksheets[0].Cells[9, 5].PutValue(ClassSat.Level3Count24);
            //ScoreWorkBook.Worksheets[0].Cells[9, 7].PutValue(ClassSat.Level3Count34);
            //ScoreWorkBook.Worksheets[0].Cells[9, 9].PutValue(ClassSat.Level3Count44);
            //ScoreWorkBook.Worksheets[0].Cells[9, 11].PutValue(ClassSat.Level3Count54);
            //ScoreWorkBook.Worksheets[0].Cells[9, 13].PutValue(ClassSat.Level3Count55);

            ////四年級班級數總計
            //ScoreWorkBook.Worksheets[0].Cells[10, 5].PutValue(ClassSat.Level4Count24);
            //ScoreWorkBook.Worksheets[0].Cells[10, 7].PutValue(ClassSat.Level4Count34);
            //ScoreWorkBook.Worksheets[0].Cells[10, 9].PutValue(ClassSat.Level4Count44);
            //ScoreWorkBook.Worksheets[0].Cells[10, 11].PutValue(ClassSat.Level4Count54);
            //ScoreWorkBook.Worksheets[0].Cells[10, 13].PutValue(ClassSat.Level4Count55);

            //一年級學生數總計
            ScoreWorkBook.Worksheets[0].Cells[9, 5].PutValue(ClassSat.SR15C5);
            ScoreWorkBook.Worksheets[0].Cells[9, 6].PutValue(ClassSat.SR15C6);
            ScoreWorkBook.Worksheets[0].Cells[9, 7].PutValue(ClassSat.SR15C7);
            ScoreWorkBook.Worksheets[0].Cells[9, 8].PutValue(ClassSat.SR15C8);
            ScoreWorkBook.Worksheets[0].Cells[9, 9].PutValue(ClassSat.SR15C9);
            ScoreWorkBook.Worksheets[0].Cells[9, 10].PutValue(ClassSat.SR15C10);
            ScoreWorkBook.Worksheets[0].Cells[9, 11].PutValue(ClassSat.SR15C11);
            ScoreWorkBook.Worksheets[0].Cells[9, 12].PutValue(ClassSat.SR15C12);
            ScoreWorkBook.Worksheets[0].Cells[9, 13].PutValue(ClassSat.SR15C13);
            ScoreWorkBook.Worksheets[0].Cells[9, 14].PutValue(ClassSat.SR15C14);

            ScoreWorkBook.Worksheets[0].Cells[10, 5].PutValue(ClassSat.SR16C5);
            ScoreWorkBook.Worksheets[0].Cells[10, 6].PutValue(ClassSat.SR16C6);
            ScoreWorkBook.Worksheets[0].Cells[10, 7].PutValue(ClassSat.SR16C7);
            ScoreWorkBook.Worksheets[0].Cells[10, 8].PutValue(ClassSat.SR16C8);
            ScoreWorkBook.Worksheets[0].Cells[10, 9].PutValue(ClassSat.SR16C9);
            ScoreWorkBook.Worksheets[0].Cells[10, 10].PutValue(ClassSat.SR16C10);
            ScoreWorkBook.Worksheets[0].Cells[10, 11].PutValue(ClassSat.SR16C11);
            ScoreWorkBook.Worksheets[0].Cells[10, 12].PutValue(ClassSat.SR16C12);
            ScoreWorkBook.Worksheets[0].Cells[10, 13].PutValue(ClassSat.SR16C13);
            ScoreWorkBook.Worksheets[0].Cells[10, 14].PutValue(ClassSat.SR16C14);

            //二年級學生數總計
            ScoreWorkBook.Worksheets[0].Cells[11, 5].PutValue(ClassSat.SR17C5);
            ScoreWorkBook.Worksheets[0].Cells[11, 6].PutValue(ClassSat.SR17C6);
            ScoreWorkBook.Worksheets[0].Cells[11, 7].PutValue(ClassSat.SR17C7);
            ScoreWorkBook.Worksheets[0].Cells[11, 8].PutValue(ClassSat.SR17C8);
            ScoreWorkBook.Worksheets[0].Cells[11, 9].PutValue(ClassSat.SR17C9);
            ScoreWorkBook.Worksheets[0].Cells[11, 10].PutValue(ClassSat.SR17C10);
            ScoreWorkBook.Worksheets[0].Cells[11, 11].PutValue(ClassSat.SR17C11);
            ScoreWorkBook.Worksheets[0].Cells[11, 12].PutValue(ClassSat.SR17C12);
            ScoreWorkBook.Worksheets[0].Cells[11, 13].PutValue(ClassSat.SR17C13);
            ScoreWorkBook.Worksheets[0].Cells[11, 14].PutValue(ClassSat.SR17C14);


            ScoreWorkBook.Worksheets[0].Cells[12, 5].PutValue(ClassSat.SR18C5);
            ScoreWorkBook.Worksheets[0].Cells[12, 6].PutValue(ClassSat.SR18C6);
            ScoreWorkBook.Worksheets[0].Cells[12, 7].PutValue(ClassSat.SR18C7);
            ScoreWorkBook.Worksheets[0].Cells[12, 8].PutValue(ClassSat.SR18C8);
            ScoreWorkBook.Worksheets[0].Cells[12, 9].PutValue(ClassSat.SR18C9);
            ScoreWorkBook.Worksheets[0].Cells[12, 10].PutValue(ClassSat.SR18C10);
            ScoreWorkBook.Worksheets[0].Cells[12, 11].PutValue(ClassSat.SR18C11);
            ScoreWorkBook.Worksheets[0].Cells[12, 12].PutValue(ClassSat.SR18C12);
            ScoreWorkBook.Worksheets[0].Cells[12, 13].PutValue(ClassSat.SR18C13);
            ScoreWorkBook.Worksheets[0].Cells[12, 14].PutValue(ClassSat.SR18C14);

            //三年級學生數總計
            ScoreWorkBook.Worksheets[0].Cells[13, 5].PutValue(ClassSat.SR19C5);
            ScoreWorkBook.Worksheets[0].Cells[13, 6].PutValue(ClassSat.SR19C6);
            ScoreWorkBook.Worksheets[0].Cells[13, 7].PutValue(ClassSat.SR19C7);
            ScoreWorkBook.Worksheets[0].Cells[13, 8].PutValue(ClassSat.SR19C8);
            ScoreWorkBook.Worksheets[0].Cells[13, 9].PutValue(ClassSat.SR19C9);
            ScoreWorkBook.Worksheets[0].Cells[13, 10].PutValue(ClassSat.SR19C10);
            ScoreWorkBook.Worksheets[0].Cells[13, 11].PutValue(ClassSat.SR19C11);
            ScoreWorkBook.Worksheets[0].Cells[13, 12].PutValue(ClassSat.SR19C12);
            ScoreWorkBook.Worksheets[0].Cells[13, 13].PutValue(ClassSat.SR19C13);
            ScoreWorkBook.Worksheets[0].Cells[13, 14].PutValue(ClassSat.SR19C14);

            ScoreWorkBook.Worksheets[0].Cells[14, 5].PutValue(ClassSat.SR20C5);
            ScoreWorkBook.Worksheets[0].Cells[14, 6].PutValue(ClassSat.SR20C6);
            ScoreWorkBook.Worksheets[0].Cells[14, 7].PutValue(ClassSat.SR20C7);
            ScoreWorkBook.Worksheets[0].Cells[14, 8].PutValue(ClassSat.SR20C8);
            ScoreWorkBook.Worksheets[0].Cells[14, 9].PutValue(ClassSat.SR20C9);
            ScoreWorkBook.Worksheets[0].Cells[14, 10].PutValue(ClassSat.SR20C10);
            ScoreWorkBook.Worksheets[0].Cells[14, 11].PutValue(ClassSat.SR20C11);
            ScoreWorkBook.Worksheets[0].Cells[14, 12].PutValue(ClassSat.SR20C12);
            ScoreWorkBook.Worksheets[0].Cells[14, 13].PutValue(ClassSat.SR20C13);
            ScoreWorkBook.Worksheets[0].Cells[14, 14].PutValue(ClassSat.SR20C14);

            ////四年級學生數總計
            //ScoreWorkBook.Worksheets[0].Cells[12, 5].PutValue(ClassSat.SR21C5);
            //ScoreWorkBook.Worksheets[0].Cells[12, 6].PutValue(ClassSat.SR21C6);
            //ScoreWorkBook.Worksheets[0].Cells[12, 7].PutValue(ClassSat.SR21C7);
            //ScoreWorkBook.Worksheets[0].Cells[12, 8].PutValue(ClassSat.SR21C8);
            //ScoreWorkBook.Worksheets[0].Cells[12, 9].PutValue(ClassSat.SR21C9);
            //ScoreWorkBook.Worksheets[0].Cells[12, 10].PutValue(ClassSat.SR21C10);
            //ScoreWorkBook.Worksheets[0].Cells[12, 11].PutValue(ClassSat.SR21C11);
            //ScoreWorkBook.Worksheets[0].Cells[12, 12].PutValue(ClassSat.SR21C12);
            //ScoreWorkBook.Worksheets[0].Cells[12, 13].PutValue(ClassSat.SR21C13);
            //ScoreWorkBook.Worksheets[0].Cells[12, 14].PutValue(ClassSat.SR21C14);

            //ScoreWorkBook.Worksheets[0].Cells[13, 5].PutValue(ClassSat.SR22C5);
            //ScoreWorkBook.Worksheets[0].Cells[13, 6].PutValue(ClassSat.SR22C6);
            //ScoreWorkBook.Worksheets[0].Cells[13, 7].PutValue(ClassSat.SR22C7);
            //ScoreWorkBook.Worksheets[0].Cells[13, 8].PutValue(ClassSat.SR22C8);
            //ScoreWorkBook.Worksheets[0].Cells[13, 9].PutValue(ClassSat.SR22C9);
            //ScoreWorkBook.Worksheets[0].Cells[13, 10].PutValue(ClassSat.SR22C10);
            //ScoreWorkBook.Worksheets[0].Cells[13, 11].PutValue(ClassSat.SR22C11);
            //ScoreWorkBook.Worksheets[0].Cells[13, 12].PutValue(ClassSat.SR22C12);
            //ScoreWorkBook.Worksheets[0].Cells[13, 13].PutValue(ClassSat.SR22C13);
            //ScoreWorkBook.Worksheets[0].Cells[13, 14].PutValue(ClassSat.SR22C14);
            //延修生學生數總計
            ScoreWorkBook.Worksheets[0].Cells[15, 5].PutValue(ClassSat.SR23C5);
            ScoreWorkBook.Worksheets[0].Cells[15, 6].PutValue(ClassSat.SR23C6);
            ScoreWorkBook.Worksheets[0].Cells[15, 7].PutValue(ClassSat.SR23C7);
            ScoreWorkBook.Worksheets[0].Cells[15, 8].PutValue(ClassSat.SR23C8);
            ScoreWorkBook.Worksheets[0].Cells[15, 9].PutValue(ClassSat.SR23C9);
            ScoreWorkBook.Worksheets[0].Cells[15, 10].PutValue(ClassSat.SR23C10);
            ScoreWorkBook.Worksheets[0].Cells[15, 11].PutValue(ClassSat.SR23C11);
            ScoreWorkBook.Worksheets[0].Cells[15, 12].PutValue(ClassSat.SR23C12);
            ScoreWorkBook.Worksheets[0].Cells[15, 13].PutValue(ClassSat.SR23C13);
            ScoreWorkBook.Worksheets[0].Cells[15, 14].PutValue(ClassSat.SR23C14);

            ScoreWorkBook.Worksheets[0].Cells[16, 5].PutValue(ClassSat.SR24C5);
            ScoreWorkBook.Worksheets[0].Cells[16, 6].PutValue(ClassSat.SR24C6);
            ScoreWorkBook.Worksheets[0].Cells[16, 7].PutValue(ClassSat.SR24C7);
            ScoreWorkBook.Worksheets[0].Cells[16, 8].PutValue(ClassSat.SR24C8);
            ScoreWorkBook.Worksheets[0].Cells[16, 9].PutValue(ClassSat.SR24C9);
            ScoreWorkBook.Worksheets[0].Cells[16, 10].PutValue(ClassSat.SR24C10);
            ScoreWorkBook.Worksheets[0].Cells[16, 11].PutValue(ClassSat.SR24C11);
            ScoreWorkBook.Worksheets[0].Cells[16, 12].PutValue(ClassSat.SR24C12);
            ScoreWorkBook.Worksheets[0].Cells[16, 13].PutValue(ClassSat.SR24C13);
            ScoreWorkBook.Worksheets[0].Cells[16, 14].PutValue(ClassSat.SR24C14);

            #endregion




            try
            {

                if (Program.ErrorList.Count > 0)
                {
                    int rowIdx = 1;
                    foreach (string str in Program.ErrorList)
                    {
                        ScoreWorkBook.Worksheets[1].Cells[rowIdx, 0].PutValue(str);
                        rowIdx++;
                    }
                    System.Windows.Forms.MessageBox.Show("產生過程有發生問題，請到工作表Error檢視。");
                }
                if (Program.ErrorList.Count == 0)
                    ScoreWorkBook.Worksheets.RemoveAt(1);
                ScoreWorkBook.Save(filename);
                System.Diagnostics.Process.Start(filename);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PrintSet_Load(object sender, EventArgs e)
        {
            List<SHDeptGroupRecord> lstdeptgroup = new List<SHDeptGroupRecord>();
            lstdeptgroup = SHDeptGroup.SelectAll();
            lstCourseKind.Items.Clear();
            //var newItem = new ComboBoxItem();
            foreach (SHDeptGroupRecord branch in lstdeptgroup)
            {
                ListViewItem lvi = new ListViewItem(branch.Name);
                lvi.SubItems.Add(branch.ID);

                lstCourseKind.Items.Add(lvi);

            }
            btnPrint.Enabled = false;
        }

        private void lstCourseKind_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;
            txtTitle.Text = "";
            for (int i = 0; i < lstCourseKind.Items.Count; i++)
                if (lstCourseKind.Items[i].Checked == true)
                {
                    btnPrint.Enabled = true;
                    txtTitle.Text = lstCourseKind.Items[i].SubItems[0].Text;
                    break;
                }
            
            
        }

        private void lstCourseKind_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            btnPrint.Enabled = false;
            txtTitle.Text = "";
            for (int i = 0; i < lstCourseKind.Items.Count; i++)
                if (lstCourseKind.Items[i].Checked == true)
                {
                    btnPrint.Enabled = true;
                    txtTitle.Text = lstCourseKind.Items[i].SubItems[0].Text;
                    break;
                }

        }

        private void txtTitle_TextChanged(object sender, EventArgs e)
        {
           
            btnPrint.Enabled = false;
            for (int i = 0; i < lstCourseKind.Items.Count; i++)
                if (lstCourseKind.Items[i].Checked == true && txtTitle.Text != "")
                {
                    btnPrint.Enabled = true;                    
                    break;
                }
        }
    }
}