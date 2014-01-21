using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace UpdateRecordReport
{
    public partial class Form1 : BaseForm
    {
        //學年度學期
        int _SchoolYear, _Semester;

        //學生物件清單
        List<Studentobj> _CleanList, _ErrorList;

        //異動代碼對照表
        Dictionary<string, string> _UpdateCode;

        //背景模式
        BackgroundWorker _BGW;

        //工作表
        Workbook _WK;

        //各科別清單
        List<Studentobj> 普通科, 綜合高中科, 職業科;

        //MappingTable
        Dictionary<string, List<string>> _MappingData;

        //Cloumn1預設選項
        List<string> _DefaultItem;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //資料初始化
            SetData();

            //取得預設學年度學期
            int schoolYear = TryParse(K12.Data.School.DefaultSchoolYear);
            int semester = TryParse(K12.Data.School.DefaultSemester);

            //cboSchoolYear選單
            for (int i = -2; i < 3; i++)
            {
                cboSchoolYear.Items.Add(schoolYear + i);
            }

            //cboSemester選單
            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");

            //cboSchoolYear cboSemester預設顯示
            cboSchoolYear.Text = schoolYear.ToString();
            cboSemester.Text = semester.ToString();

            //Column1選單
            foreach (string s in _DefaultItem)
            {
                this.Column1.Items.Add(s);
            }

            //Column2選單
            foreach (string s in _UpdateCode.Keys)
            {
                this.Column2.Items.Add(s);
            }

            //讀取上次設定
            LoadLastRecord();

            //_BGW
            _BGW = new BackgroundWorker();
            _BGW.DoWork += new DoWorkEventHandler(_BGW_DoWork);
            _BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGW_RunWorkerCompleted);
        }

        private void SetData()
        {
            #region Column1選單項目
            _DefaultItem = new List<string>();
            _DefaultItem.Add("轉出:遷居");
            _DefaultItem.Add("轉出:家長調職");
            _DefaultItem.Add("轉出:改變環境");
            _DefaultItem.Add("轉出:輔導轉學");
            _DefaultItem.Add("轉出:其他");

            _DefaultItem.Add("退學:自動退學");
            _DefaultItem.Add("退學:休學期滿");
            _DefaultItem.Add("退學:未達畢業標準");
            _DefaultItem.Add("退學:其他");

            _DefaultItem.Add("休學:因病");
            _DefaultItem.Add("休學:志趣不合");
            _DefaultItem.Add("休學:經濟困難");
            _DefaultItem.Add("休學:兵役");
            _DefaultItem.Add("休學:出國");
            _DefaultItem.Add("休學:其他");

            _DefaultItem.Add("復學生");

            _DefaultItem.Add("轉入:他校轉入");
            _DefaultItem.Add("轉入:本校不同學制轉入");

            _DefaultItem.Add("死亡");
            _DefaultItem.Add("輔導延修");
            #endregion

            //建立異動代碼對照表
            _UpdateCode = new Dictionary<string, string>();
            //讀取XML
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(Properties.Resources.UpdateCode_SHD);

            foreach (XmlElement Xelem in doc.DocumentElement)
            {
                string code = "";
                string reason = "";
                bool mustAdd = false;

                foreach (XmlElement elem in Xelem.ChildNodes)
                {
                    if (elem.Name == "代號")
                    {
                        code = elem.InnerText;
                    }

                    if (elem.Name == "原因及事項")
                    {
                        reason = elem.InnerText;
                    }

                    //只加入學籍異動類
                    if (elem.Name == "分類")
                    {
                        if (elem.InnerText == "學籍異動") mustAdd = true;
                    }
                }

                if (!_UpdateCode.ContainsKey(code) && mustAdd)
                {
                    _UpdateCode.Add(code + ":" + reason, code);
                }
            }
        }

        private void _BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //解除表單元件封鎖
            FormControlEnabled(true);

            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 高中職學校學生異動報告 已完成");

            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "高中職學校學生異動報告.xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _WK.Save(sd.FileName);
                    if (_ErrorList.Count > 0)
                    {
                        MessageBox.Show("發現" + _ErrorList.Count + "筆異常資料未列入統計\r\n詳細資料請確認報表中的[異常資料表]");
                    }
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    this.Enabled = true;
                    return;
                }
            }
        }

        private void _BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //建立學生物件字典
            Dictionary<string, Studentobj> studentObjDic = new Dictionary<string, Studentobj>();

            //SQL查詢資料
            QueryHelper _Q = new QueryHelper();
            StringBuilder sb = new StringBuilder();
            sb.Append("select update_record.ref_student_id,update_record.update_code,student.name,student.gender,student.student_number,class.grade_year,dept.name as dept_name from update_record ");
            sb.Append("left join student on update_record.ref_student_id = student.id ");
            sb.Append("left join class on student.ref_class_id = class.id ");
            sb.Append("left join dept on class.ref_dept_id = dept.id ");
            sb.Append(string.Format("where update_record.school_year='{0}' and update_record.semester='{1}' ", _SchoolYear, _Semester));

            DataTable dt = _Q.Select(sb.ToString());

            foreach (DataRow row in dt.Rows)
            {
                string id = row["ref_student_id"].ToString();
                string name = row["name"].ToString();
                string gender = row["gender"].ToString();
                string student_number = row["student_number"].ToString();
                string grade_year = row["grade_year"].ToString();
                string dept_name = row["dept_name"].ToString();
                string code = row["update_code"].ToString();

                //字典不存在學生ID就新增
                if (!studentObjDic.ContainsKey(id))
                {
                    studentObjDic.Add(id, new Studentobj(id, name, gender, student_number, grade_year, dept_name));
                }

                //該學生物件的CodeList不存在此cdoe就加入
                if (!studentObjDic[id].CodeList.Contains(code))
                {
                    studentObjDic[id].CodeList.Add(code);
                }
            }

            _CleanList = new List<Studentobj>();
            _ErrorList = new List<Studentobj>();

            foreach (KeyValuePair<string, Studentobj> k in studentObjDic)
            {
                //排除畢業代碼501
                if (k.Value.CodeList.Contains("501")) continue;

                if(k.Value.Gender == "1" || k.Value.Gender == "0")
                {
                    _CleanList.Add(k.Value);
                }
                else
                {
                    _ErrorList.Add(k.Value);
                }
            }

            普通科 = getStudentListByDept("普通科");
            綜合高中科 = getStudentListByDept("綜合高中科");
            職業科 = getStudentListByDept("職業科");

            Export();
        }

        private void Export()
        {
            _WK = new Workbook();
            Workbook wk2 = new Workbook();
            wk2.Open(new MemoryStream(Properties.Resources.Template)); //開啟範本文件
            _WK.Worksheets[0].Copy(wk2.Worksheets[0]); //複製範本文件
            Worksheet ws = _WK.Worksheets[0];
            ws.Name = "普通科";
            Cells cs = ws.Cells;
            int index = 8;
            //需要跳下一行的行數
            List<int> nextRow = new List<int>(new int[]{13,18,26}) ;

            Dictionary<string, List<Studentobj>> dica = getSortDic(普通科);
            foreach(KeyValuePair<string, List<Studentobj>> k in dica)
            {
                cs[index, 1].PutValue(k.Value.Count);
                cs[index, 2].PutValue(getStudentCount(k.Value,"0","1"));
                cs[index, 3].PutValue(getStudentCount(k.Value, "0", "0"));
                cs[index, 4].PutValue(getStudentCount(k.Value, "1", "1"));
                cs[index, 5].PutValue(getStudentCount(k.Value, "1", "0"));
                cs[index, 6].PutValue(getStudentCount(k.Value, "2", "1"));
                cs[index, 7].PutValue(getStudentCount(k.Value, "2", "0"));
                cs[index, 8].PutValue(getStudentCount(k.Value, "3", "1"));
                cs[index, 9].PutValue(getStudentCount(k.Value, "3", "0"));
                cs[index, 10].PutValue(getStudentCount(k.Value, "4", "1"));
                cs[index, 11].PutValue(getStudentCount(k.Value, "4", "0"));
                cs[index, 12].PutValue(getDelayCount(k.Value, "1"));
                cs[index, 13].PutValue(getDelayCount(k.Value, "0"));
                index++;
                if (nextRow.Contains(index)) index++;
            }

            _WK.Worksheets.Add();
            _WK.Worksheets[1].Copy(wk2.Worksheets[0]); //複製範本文件
            ws = _WK.Worksheets[1];
            ws.Name = "綜合高中科";
            cs = ws.Cells;
            index = 8;

            Dictionary<string, List<Studentobj>> dicb = getSortDic(綜合高中科);
            foreach (KeyValuePair<string, List<Studentobj>> k in dicb)
            {
                cs[index, 1].PutValue(k.Value.Count);
                cs[index, 2].PutValue(getStudentCount(k.Value, "0", "1"));
                cs[index, 3].PutValue(getStudentCount(k.Value, "0", "0"));
                cs[index, 4].PutValue(getStudentCount(k.Value, "1", "1"));
                cs[index, 5].PutValue(getStudentCount(k.Value, "1", "0"));
                cs[index, 6].PutValue(getStudentCount(k.Value, "2", "1"));
                cs[index, 7].PutValue(getStudentCount(k.Value, "2", "0"));
                cs[index, 8].PutValue(getStudentCount(k.Value, "3", "1"));
                cs[index, 9].PutValue(getStudentCount(k.Value, "3", "0"));
                cs[index, 10].PutValue(getStudentCount(k.Value, "4", "1"));
                cs[index, 11].PutValue(getStudentCount(k.Value, "4", "0"));
                cs[index, 12].PutValue(getDelayCount(k.Value, "1"));
                cs[index, 13].PutValue(getDelayCount(k.Value, "0"));
                index++;
                if (nextRow.Contains(index)) index++;
            }

            _WK.Worksheets.Add();
            _WK.Worksheets[2].Copy(wk2.Worksheets[0]); //複製範本文件
            ws = _WK.Worksheets[2];
            ws.Name = "職業科";
            cs = ws.Cells;
            index = 8;

            Dictionary<string, List<Studentobj>> dicc = getSortDic(職業科);
            foreach (KeyValuePair<string, List<Studentobj>> k in dicc)
            {
                cs[index, 1].PutValue(k.Value.Count);
                cs[index, 2].PutValue(getStudentCount(k.Value, "0", "1"));
                cs[index, 3].PutValue(getStudentCount(k.Value, "0", "0"));
                cs[index, 4].PutValue(getStudentCount(k.Value, "1", "1"));
                cs[index, 5].PutValue(getStudentCount(k.Value, "1", "0"));
                cs[index, 6].PutValue(getStudentCount(k.Value, "2", "1"));
                cs[index, 7].PutValue(getStudentCount(k.Value, "2", "0"));
                cs[index, 8].PutValue(getStudentCount(k.Value, "3", "1"));
                cs[index, 9].PutValue(getStudentCount(k.Value, "3", "0"));
                cs[index, 10].PutValue(getStudentCount(k.Value, "4", "1"));
                cs[index, 11].PutValue(getStudentCount(k.Value, "4", "0"));
                cs[index, 12].PutValue(getDelayCount(k.Value, "1"));
                cs[index, 13].PutValue(getDelayCount(k.Value, "0"));
                index++;
                if (nextRow.Contains(index)) index++;
            }

            _WK.Worksheets.Add();
            ws = _WK.Worksheets[3];
            ws.Name = "異常資料表";
            cs = ws.Cells;
            cs["A1"].PutValue("學號");
            cs["B1"].PutValue("姓名");
            cs["C1"].PutValue("性別");
            cs["D1"].PutValue("年級");
            cs["E1"].PutValue("科別名稱");
            index = 1;
            foreach (Studentobj s in _ErrorList)
            {
                cs[index, 0].PutValue(s.Student_number);
                cs[index, 1].PutValue(s.Name);
                cs[index, 2].PutValue(s.Gender);
                cs[index, 3].PutValue(s.Grade);
                cs[index, 4].PutValue(s.Dept);
                index++;
            }
        }

        private List<Studentobj> getStudentListByDept(String dept)
        {
            List<Studentobj> list = new List<Studentobj>();

            switch (dept)
            {
                case "職業科":
                    foreach (Studentobj student in _CleanList)
                    {
                        if (!student.Dept.Contains("普通科") && !student.Dept.Contains("綜合高中科"))
                        {
                            list.Add(student);
                        }
                    }
                    break;

                default:
                    foreach (Studentobj student in _CleanList)
                    {
                        if (student.Dept.Contains(dept))
                        {
                            list.Add(student);
                        }
                    }
                    break;
            }
            return list;
        }

        private Dictionary<string, List<Studentobj>> getSortDic(List<Studentobj> list)
        {
            Dictionary<string, List<Studentobj>> dic = new Dictionary<string, List<Studentobj>>();

            foreach (string key in _MappingData.Keys)
            {
                //建立預設的鍵值 from _MappingData
                dic.Add(key, new List<Studentobj>());

                //循環每個項目對應的代碼
                foreach (string code in _MappingData[key])
                {
                    //循環每個學生找代碼
                    foreach (Studentobj student in list)
                    {
                        //有代碼者加入字典
                        if (student.CodeList.Contains(code))
                        {
                            dic[key].Add(student);
                        }
                    }
                }
            }

            return dic;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                //儲存畫面紀錄
                SaveMappingRecord();
                //學年度學期
                _SchoolYear = TryParse(cboSchoolYear.Text);
                _Semester = TryParse(cboSemester.Text);
                //讀取畫面資料
                ReadMappingData();
                //封鎖表單元件
                FormControlEnabled(false);
                //啟動背景執行
                _BGW.RunWorkerAsync();
            }
            catch
            {
                MessageBox.Show("網路或資料庫異常,請稍後再試...");
                this.btnOk.Enabled = true;
                this.linkLabel1.Enabled = true;
                this.dataGridViewX1.Enabled = true;
            }
        }

        //簡單字串轉數字
        private int TryParse(string s)
        {
            int i = 0;
            try
            {
                i = int.Parse(s);
            }
            catch
            {
                i = 0;
            }

            return i;
        }

        public void SaveMappingRecord() //儲存上次Mapping紀錄
        {
            AccessHelper _A = new AccessHelper();
            List<UpdateRecordReportUDT> UDTlist = _A.Select<UpdateRecordReportUDT>();
            _A.DeletedValues(UDTlist); //清除UDT資料

            UDTlist = new List<UpdateRecordReportUDT>(); //清空UDTlist
            foreach (DataGridViewRow row in dataGridViewX1.Rows) //取得DataDataGridViewRow資料
            {
                if (row.Cells[0].Value == null) //遇到空白的Target即跳到下個loop
                {
                    continue;
                }

                String target = row.Cells[0].Value.ToString();
                String source = "";
                if (row.Cells[1].Value != null) { source = row.Cells[1].Value.ToString(); }

                UpdateRecordReportUDT obj = new UpdateRecordReportUDT();
                obj.Target = target;
                obj.Source = source;
                UDTlist.Add(obj);
            }

            _A.InsertValues(UDTlist); //回存到UDT
        }

        void ReadMappingData() //讀取DataGridView資料
        {
            _MappingData = new Dictionary<string, List<string>>();

            //建立預設建值
            foreach(string item in _DefaultItem)
            {
                _MappingData.Add(item, new List<string>());
            }

            //讀取畫面資料
            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if (r.Cells[0].Value != null && r.Cells[1].Value != null)  //欄位有空值跳下一行
                {
                    string target = r.Cells[0].Value.ToString();
                    string source = r.Cells[1].Value.ToString();

                    //取得異動代碼
                    string code = _UpdateCode.ContainsKey(source) ? _UpdateCode[source] : "";

                    if (target != "" && code != "")
                    {
                        if (!_MappingData[target].Contains(code))
                        {
                            _MappingData[target].Add(code);
                        }
                    }
                }
            }
        }

        public void LoadLastRecord() //讀取上次Mapping設定
        {
            AccessHelper _A = new AccessHelper();
            List<UpdateRecordReportUDT> UDTlist = _A.Select<UpdateRecordReportUDT>(); //檢查UDT並回傳資料
            DataGridViewRow row;
            if (UDTlist.Count > 0) //UDT內有設定才做讀取
            {
                for (int i = 0; i < UDTlist.Count; i++)
                {
                    row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = UDTlist[i].Target;
                    if (UDTlist[i].Source == "")
                    {
                        row.Cells[1].Value = null;
                    }
                    else
                    {
                        row.Cells[1].Value = UDTlist[i].Source;
                    }
                    dataGridViewX1.Rows.Add(row);
                }
            }
            else
            {
                //UDT無資料則提供預設標記
                for (int i = 0; i < Column1.Items.Count; i++)
                {
                    row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = Column1.Items[i];
                    dataGridViewX1.Rows.Add(row);
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                AccessHelper _A = new AccessHelper();
                List<UpdateRecordReportUDT> UDTlist = _A.Select<UpdateRecordReportUDT>();
                _A.DeletedValues(UDTlist); //清除UDT資料
                dataGridViewX1.Rows.Clear();  //清除datagridview資料
                LoadLastRecord(); //再次讀入Mapping設定
            }
            catch
            {
                MessageBox.Show("網路或資料庫異常,請稍後再試...");
            }
        }

        //取得指定年級性別的一般生總數
        private int getStudentCount(List<Studentobj> list, String grade, String gender)
        {
            int count = 0;
            switch (grade)
            {
                case "0": //不指定年級
                    foreach (Studentobj s in list)
                    {
                        if (s.Gender == gender)
                        {
                            count++;
                        }
                    }
                    break;
                default:
                    foreach (Studentobj s in list)
                    {
                        if (s.Grade == grade && s.Gender == gender)
                        {
                            count++;
                        }
                    }
                    break;
            }
            return count;
        }

        //取得延修身分的學生數量
        private int getDelayCount(List<Studentobj> list,string gender)
        {
            int count = 0;
            //延修代碼
            List<string> delayCodes = new List<string>(new string[]{"235","236"});

            foreach(Studentobj student in list)
            {
                foreach(string code in delayCodes)
                {
                    if (student.CodeList.Contains(code) && student.Gender == gender) count++;
                }
            }

            return count;
        }

        //封鎖表單元件
        private void FormControlEnabled(bool b)
        {
            this.btnOk.Enabled = b;
            this.cboSchoolYear.Enabled = b;
            this.cboSemester.Enabled = b;
            this.dataGridViewX1.Enabled = b;
            this.linkLabel1.Enabled = b;
        }
    }
}
