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
        List<RecordObj> _CleanList, _ErrorList;

        //異動代碼對照表
        Dictionary<string, string> _QueryUpdateCodes;

        //Column2預設選項
        Dictionary<string, string> _UpdateItems;
        //Column1預設選項
        Dictionary<string, List<string>> _DefaultItem;
        //MappingTable
        Dictionary<string, List<string>> _MappingData;

        //背景模式
        BackgroundWorker _BGW;

        //工作表
        Workbook _WK;

        //各科別清單
        //List<RecordObj> 普通科, 綜合高中科, 職業科;
        string Public_BranchID;
        string Public_BranchName;
        
        public Form1(string BranchID, string BranchName)
        {
            InitializeComponent();
            Public_BranchID = BranchID;
            Public_BranchName = BranchName;
            
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
            foreach (string s in _DefaultItem.Keys)
            {
                this.Column1.Items.Add(s);
            }

            //Column2選單
            foreach (string s in _UpdateItems.Keys)
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
            _DefaultItem = new Dictionary<string, List<string>>();
            _DefaultItem.Add("轉出:遷居",new List<string>(new string[]{"311"}));
            _DefaultItem.Add("轉出:家長調職", new List<string>(new string[] { "312" }));
            _DefaultItem.Add("轉出:改變環境", new List<string>(new string[] { "313" }));
            //_DefaultItem.Add("轉出:輔導轉學", new List<string>(new string[] { "315" }));
            _DefaultItem.Add("轉出:其他", new List<string>(new string[] { "314","316" }));

            _DefaultItem.Add("放棄、廢止、註銷學籍:主動辦理放棄學籍", new List<string>(new string[] { "367", "369", "378","379" }));//379 不計人數
            _DefaultItem.Add("放棄、廢止、註銷學籍:因休學期滿而放棄、廢止學籍", new List<string>(new string[] { "380","381" })); //380 不計人數
            _DefaultItem.Add("放棄、廢止、註銷學籍:其他(含註銷學籍)", new List<string>(new string[] { "374", "375" }));  //375 不計人數
            //_DefaultItem.Add("退學:自動退學", new List<string>(new string[] { "321" }));
            //_DefaultItem.Add("退學:休學期滿", new List<string>(new string[] { "323" }));
            //_DefaultItem.Add("退學:未達畢業標準",new List<string>());
            //_DefaultItem.Add("退學:其他", new List<string>(new string[] { "325","326" }));

            _DefaultItem.Add("休學:因病", new List<string>(new string[] { "341" }));
            _DefaultItem.Add("休學:志趣不合", new List<string>(new string[] { "342" }));
            _DefaultItem.Add("休學:經濟困難", new List<string>(new string[] { "343" }));
            _DefaultItem.Add("休學:兵役", new List<string>(new string[] { "345" }));
            _DefaultItem.Add("休學:出國", new List<string>(new string[] { "348" }));
            _DefaultItem.Add("休學:缺曠課過多", new List<string>(new string[] { "344" }));
            _DefaultItem.Add("休學:其他", new List<string>(new string[] { "346","347","349" }));

            _DefaultItem.Add("復學生", new List<string>(new string[] { "221", "222", "223", "224", "225", "226", "237", "238", "239", "240", }));

            _DefaultItem.Add("轉入:他校轉入", new List<string>(new string[] { "111", "112", "113", "114", "115", "121", "122", "123", "124", }));
            _DefaultItem.Add("轉入:本校不同學制轉入",new List<string>());

            _DefaultItem.Add("死亡", new List<string>(new string[] { "361" }));
            _DefaultItem.Add("輔導延修", new List<string>(new string[] { "364" }));
            _DefaultItem.Add("修業年限期滿", new List<string>(new string[] { "365","372" }));
            //_DefaultItem.Add("未達畢業標準", new List<string>(new string[] { "366" }));

            _DefaultItem.Add("未達畢業標準(指德性評量) ", new List<string>(new string[] { "366" }));
            #endregion

            //建立異動代碼對照表
            _UpdateItems = new Dictionary<string, string>();
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

                    //只加入學籍異動類和轉入異動
                    if (elem.Name == "分類")
                    {
                        if (elem.InnerText == "學籍異動" || elem.InnerText == "轉入異動")
                        {
                            //延修代碼不可為選項
                            if (code != "235" && code != "236" && code != "243" && code != "244") mustAdd = true;
                        } 
                    }
                }

                if (!_UpdateItems.ContainsKey(code) && mustAdd)
                {
                    _UpdateItems.Add(code + ":" + reason, code);
                }
            }

            //查詢代碼用字典
            _QueryUpdateCodes = new Dictionary<string, string>();
            foreach (KeyValuePair<string, string> k in _UpdateItems)
            {
                if (!_QueryUpdateCodes.ContainsKey(k.Value))
                    _QueryUpdateCodes.Add(k.Value, k.Key);
            }
        }

        private void _BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //解除表單元件封鎖
            FormControlEnabled(true);

            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 高中職學校學生異動報告 已完成");

            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "高中職學校學生異動報告.xlsx";
            sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
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
                    this.Close();
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
            Dictionary<string, RecordObj> RecordObjDic = new Dictionary<string, RecordObj>();

            //組SQL(強制加入501,235,236)
            string updateCode = "'501','235','236','243','244','"; //畢業,延修一,延修二
            foreach (List<string> codes in _MappingData.Values)
            {
                foreach(string s in codes)
                {
                    updateCode += s + "','";
                }
            }
            updateCode = updateCode + "'";

            //SQL查詢資料
            QueryHelper _Q = new QueryHelper();
            StringBuilder sb = new StringBuilder();

            sb.Append(@"WITH student_data AS (
                        SELECT update_record.id, update_record.ref_student_id, update_record.ss_name
                         , ss_student_number, update_record.ss_gender, update_record.ss_grade_year,update_record.ss_dept
                         , COALESCE(student.ref_dept_id,class.ref_dept_id ) AS _dept
                         , update_code, school_year,semester, student.status
                          FROM student JOIN class ON student.ref_class_id=class.id
                               LEFT JOIN update_record ON update_record.ref_student_id=student.id )
                    SELECT
                            student_data.*, dept.name AS dept_name, dept.ref_dept_group_id
                        FROM
                            student_data JOIN  dept ON student_data._dept= dept.id");
            //sb.Append("select update_record.id,ref_student_id,ss_name,ss_student_number,ss_gender,ss_grade_year,ss_dept,update_code,student.status from update_record left join student on ref_student_id = student.id left join class on student.ref_class_id=class.id left join dept on class.ref_dept_id=dept.id ");
            sb.Append(string.Format(" where school_year={0} and semester={1} and update_code in ({2}) and dept.ref_dept_group_id in ({3}) ", _SchoolYear, _Semester, updateCode,  Public_BranchID.Substring(0,Public_BranchID.Length-1)));

            DataTable dt = _Q.Select(sb.ToString());

            //紀錄有501畢業代碼紀錄的學生ID
            List<string> graduateList = new List<string>();
            //紀錄有245,236延修代碼紀錄的學生ID
            List<string> delayList = new List<string>();

            foreach (DataRow row in dt.Rows)
            {
                string uid = row["id"].ToString();
                string sid = row["ref_student_id"].ToString();
                string code = row["update_code"].ToString();
                
                //有501代碼且不存在於passList者,sid加入清單
                if (code == "501" && !graduateList.Contains(sid))
                {
                    graduateList.Add(sid);
                }

                //有235或236代碼且不存在於delayList者,sid加入清單
                if ((code == "235" || code == "236" || code == "243" || code == "244") && !delayList.Contains(sid))
                {
                   
                    delayList.Add(sid);
                }

                //字典不存在UID就新增
                if (!RecordObjDic.ContainsKey(uid))
                {
                    RecordObjDic.Add(uid, new RecordObj(row));
                }
            }

            _CleanList = new List<RecordObj>();
            _ErrorList = new List<RecordObj>();

            //過濾資料
            foreach (KeyValuePair<string, RecordObj> k in RecordObjDic)
            {
                //判斷是否繼續處理
                bool canPass = false;
                //學生系統編號存在於passList者,不做處理
                if (graduateList.Contains(k.Value.Student_id)) continue;
                //學生系統編號存在於delayList者,且代碼為235或236不做處理
                if (delayList.Contains(k.Value.Student_id))
                {
                    if (k.Value.Code == "235" || k.Value.Code == "236" || k.Value.Code == "243" || k.Value.Code == "244")
                    {
                        continue;
                    }
                    else
                    {
                        //不是的話標記為延修身分
                        k.Value.Delay = true;
                    }
                }

                //不是_MappingData有選的異動代碼,不做處理
                foreach (List<string> codes in _MappingData.Values)
                {
                    if (codes.Contains(k.Value.Code))
                    {
                        canPass = true;
                        break;
                    }
                }

                //canPass為false者跳過不處理
                if (!canPass) continue;

                // 2022-08-30 只要異動是"延修生" 就直接視為延修生
                if (k.Value.Grade == "-1")
                    k.Value.Delay = true;

                if (k.Value.Status == "2")
                    k.Value.Grade = "-1";
                //性別為男或女且有正確(1~4)年級者才處理,否則加入錯誤清單 //2022/8/26 將 "-1" 延修生也列為正確
                if ((k.Value.Gender == "1" || k.Value.Gender == "0") && (k.Value.Grade == "1" || k.Value.Grade == "2" || k.Value.Grade == "3" || k.Value.Grade == "4" || k.Value.Grade == "-1"))
                {
                    _CleanList.Add(k.Value);
                }
                else
                {
                    _ErrorList.Add(k.Value);
                }
            }

            //普通科 = getStudentListByDept("普通科");
            //綜合高中科 = getStudentListByDept("綜合高中科");
            //職業科 = getStudentListByDept("職業科");

            Export();
        }

        private void Export()
        {
            _WK = new Workbook();
            Workbook wk2 = new Workbook();
            wk2.Open(new MemoryStream(Properties.Resources.Template)); //開啟範本文件
            _WK.Worksheets[0].Copy(wk2.Worksheets[0]); //複製範本文件
            Worksheet ws = _WK.Worksheets[0];
            ws.Name = Public_BranchName;
            Cells cs = ws.Cells;
            int index = 8;
            //需要跳下一行的行數
            //List<int> nextRow = new List<int>(new int[] { 13, 18, 27 });
            List<int> nextRow = new List<int>(new int[] { 12, 16, 25 });

            Dictionary<string, List<RecordObj>> dica = getSortDic(_CleanList);
            cs[3, 5].PutValue(_SchoolYear + " 學年第 " + _Semester + "     學期");
            cs[0, 10].PutValue(K12.Data.School.ChineseName + "(教務處)");
            cs[4, 0].PutValue(K12.Data.School.Code);
            cs[4, 1].PutValue(Public_BranchName);
            cs[2, 0].PutValue("表 - 16 高級中等學校學生異動概況─"+Public_BranchName);
            foreach (KeyValuePair<string, List<RecordObj>> k in dica)
            {
                cs[index, 2].PutValue(k.Value.Count);
                cs[index, 3].PutValue(getStudentCount(k.Value, "0", "1"));
                cs[index, 4].PutValue(getStudentCount(k.Value, "0", "0"));
                cs[index, 5].PutValue(getStudentCount(k.Value, "1", "1"));
                cs[index, 6].PutValue(getStudentCount(k.Value, "1", "0"));
                cs[index, 7].PutValue(getStudentCount(k.Value, "2", "1"));
                cs[index, 8].PutValue(getStudentCount(k.Value, "2", "0"));
                cs[index, 9].PutValue(getStudentCount(k.Value, "3", "1"));
                cs[index, 10].PutValue(getStudentCount(k.Value, "3", "0"));
                //cs[index, 11].PutValue(getStudentCount(k.Value, "4", "1"));
                //cs[index, 12].PutValue(getStudentCount(k.Value, "4", "0"));
                cs[index, 11].PutValue(getDelayCount(k.Value, "1"));
                cs[index, 12].PutValue(getDelayCount(k.Value, "0"));
                index++;
                if (nextRow.Contains(index)) index++;
            }

            //_WK.Worksheets.Add();
            //_WK.Worksheets[1].Copy(wk2.Worksheets[0]); //複製範本文件
            //ws = _WK.Worksheets[1];
            //ws.Name = "綜合高中科";
            //cs = ws.Cells;
            //index = 8;

            //Dictionary<string, List<RecordObj>> dicb = getSortDic(綜合高中科);
            //cs[3, 5].PutValue(_SchoolYear + " 學年第 " + _Semester + "     學期");
            //foreach (KeyValuePair<string, List<RecordObj>> k in dicb)
            //{
            //    cs[index, 1].PutValue(k.Value.Count);
            //    cs[index, 2].PutValue(getStudentCount(k.Value, "0", "1"));
            //    cs[index, 3].PutValue(getStudentCount(k.Value, "0", "0"));
            //    cs[index, 4].PutValue(getStudentCount(k.Value, "1", "1"));
            //    cs[index, 5].PutValue(getStudentCount(k.Value, "1", "0"));
            //    cs[index, 6].PutValue(getStudentCount(k.Value, "2", "1"));
            //    cs[index, 7].PutValue(getStudentCount(k.Value, "2", "0"));
            //    cs[index, 8].PutValue(getStudentCount(k.Value, "3", "1"));
            //    cs[index, 9].PutValue(getStudentCount(k.Value, "3", "0"));
            //    cs[index, 10].PutValue(getStudentCount(k.Value, "4", "1"));
            //    cs[index, 11].PutValue(getStudentCount(k.Value, "4", "0"));
            //    cs[index, 12].PutValue(getDelayCount(k.Value, "1"));
            //    cs[index, 13].PutValue(getDelayCount(k.Value, "0"));
            //    index++;
            //    if (nextRow.Contains(index)) index++;
            //}

            //_WK.Worksheets.Add();
            //_WK.Worksheets[2].Copy(wk2.Worksheets[0]); //複製範本文件
            //ws = _WK.Worksheets[2];
            //ws.Name = "職業科";
            //cs = ws.Cells;
            //index = 8;

            //Dictionary<string, List<RecordObj>> dicc = getSortDic(職業科);
            //cs[3, 5].PutValue(_SchoolYear + " 學年第 " + _Semester + "     學期");
            //foreach (KeyValuePair<string, List<RecordObj>> k in dicc)
            //{
            //    cs[index, 1].PutValue(k.Value.Count);
            //    cs[index, 2].PutValue(getStudentCount(k.Value, "0", "1"));
            //    cs[index, 3].PutValue(getStudentCount(k.Value, "0", "0"));
            //    cs[index, 4].PutValue(getStudentCount(k.Value, "1", "1"));
            //    cs[index, 5].PutValue(getStudentCount(k.Value, "1", "0"));
            //    cs[index, 6].PutValue(getStudentCount(k.Value, "2", "1"));
            //    cs[index, 7].PutValue(getStudentCount(k.Value, "2", "0"));
            //    cs[index, 8].PutValue(getStudentCount(k.Value, "3", "1"));
            //    cs[index, 9].PutValue(getStudentCount(k.Value, "3", "0"));
            //    cs[index, 10].PutValue(getStudentCount(k.Value, "4", "1"));
            //    cs[index, 11].PutValue(getStudentCount(k.Value, "4", "0"));
            //    cs[index, 12].PutValue(getDelayCount(k.Value, "1"));
            //    cs[index, 13].PutValue(getDelayCount(k.Value, "0"));
            //    index++;
            //    if (nextRow.Contains(index)) index++;
            //}

            _WK.Worksheets.Add();
            ws = _WK.Worksheets[1];
            ws.Name = "異常資料表";
            cs = ws.Cells;
            cs["A1"].PutValue("學號");
            cs["B1"].PutValue("姓名");
            cs["C1"].PutValue("性別");
            cs["D1"].PutValue("年級");
            cs["E1"].PutValue("科別名稱");
            cs["E1"].PutValue("異動代碼");
            index = 1;
            foreach (RecordObj s in _ErrorList)
            {
                //無學號者提供系統編號
                cs[index, 0].PutValue(s.Student_number == "" ? "系統編號: " + s.Student_id : s.Student_number);
                cs[index, 1].PutValue(s.Name);
                cs[index, 2].PutValue(s.Gender);
                cs[index, 3].PutValue(s.Grade);
                cs[index, 4].PutValue(s.Dept);
                cs[index, 4].PutValue(_QueryUpdateCodes[s.Code]);
                index++;
            }
        }

        //private List<RecordObj> getStudentListByDept(String dept)
        //{
        //    List<RecordObj> list = new List<RecordObj>();

        //    switch (dept)
        //    {
        //        case "職業科":
        //            foreach (RecordObj obj in _CleanList)
        //            {
        //                if (!obj.Dept.Contains("普通科") && !obj.Dept.Contains("綜合高中科"))
        //                {
        //                    list.Add(obj);
        //                }
        //            }
        //            break;

        //        default:
        //            foreach (RecordObj obj in _CleanList)
        //            {
        //                if (obj.Dept.Contains(dept))
        //                {
        //                    list.Add(obj);
        //                }
        //            }
        //            break;
        //    }
        //    return list;
        //}

        //將學生清單分類成_MappingData的各項目
        private Dictionary<string, List<RecordObj>> getSortDic(List<RecordObj> list)
        {
            Dictionary<string, List<RecordObj>> dic = new Dictionary<string, List<RecordObj>>();

            foreach (string key in _MappingData.Keys)
            {
                //建立預設的鍵值 from _MappingData
                dic.Add(key, new List<RecordObj>());

                //循環每個項目對應的代碼
                foreach (string code in _MappingData[key])
                {
                    //循環每筆記錄找代碼
                    foreach (RecordObj obj in list)
                    {
                        //有代碼者加入字典
                        if (obj.Code == code)
                        {
                            dic[key].Add(obj);
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
            foreach (string item in _DefaultItem.Keys)
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
                    string code = _UpdateItems.ContainsKey(source) ? _UpdateItems[source] : "";

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
                foreach(KeyValuePair<string,List<string>> k in _DefaultItem)
                {
                    //若沒有預設代碼就給一個空白的row
                    if(k.Value.Count == 0)
                    {
                        row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX1);
                        row.Cells[0].Value = k.Key;
                        dataGridViewX1.Rows.Add(row);
                    }
                    else
                    {
                        //有預設代碼將代碼帶入預設選項
                        foreach (string s in k.Value)
                        {
                            row = new DataGridViewRow();
                            row.CreateCells(dataGridViewX1);
                            row.Cells[0].Value = k.Key;
                            row.Cells[1].Value = _QueryUpdateCodes[s];
                            dataGridViewX1.Rows.Add(row);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 重設
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        private int getStudentCount(List<RecordObj> list, String grade, String gender)
        {
            int count = 0;
            switch (grade)
            {
                case "0": //不指定年級
                    foreach (RecordObj s in list)
                    {
                        if (s.Gender == gender)
                        {
                            count++;
                        }
                    }
                    break;
                default:
                    foreach (RecordObj s in list)
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
        private int getDelayCount(List<RecordObj> list, string gender)
        {
            int count = 0;

            foreach (RecordObj obj in list)
            {
                if (obj.Delay && obj.Gender == gender) count++;
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
