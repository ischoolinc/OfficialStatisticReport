using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation.Controls;
using FISCA.UDT;
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

namespace myTable
{
    public partial class Form2 : BaseForm
    {
        String _SchoolYear; //當前學年度
        private BackgroundWorker _BGWClassStudentAbsenceDetail; //背景模式
        Dictionary<String, String> _column3Items; //全部類別對照表,key=TagId,value=prefix+":"+name
        Dictionary<String, List<String>> _mappingData;//mapping資料
        Filter filter;
        List<String> AboList; //原住民生清單
        String dept;
        Workbook _wk;

        public Form2()
        {
            InitializeComponent();
            Column2Prepare();
            Column3Prepare();
            SchoolYearItem();
            LoadLastRecord();
        }

        ////Column2的選單產生
        private void Column2Prepare()
        {
            List<String> prefix = new List<string>();
            List<String> name = new List<string>();
            prefix.Add("甄選入學");
            prefix.Add("申請入學");
            prefix.Add("登記分發");
            prefix.Add("直升入學");
            prefix.Add("免試入學");
            prefix.Add("其他");
            name.Add("一般生");
            name.Add("原住民生");
            name.Add("身心障礙生");
            name.Add("其他");

            foreach (String a in prefix)
            {
                foreach (String b in name)
                {
                    Column2.Items.Add(a + ":" + b);
                }
            }
        }

        //Column3的選單產生
        private void Column3Prepare()
        {
            _column3Items = new Dictionary<String, String>();
            QueryHelper _Q = new QueryHelper();

            DataTable dt = _Q.Select("select * from tag where category='Student' order by prefix,name");
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String prefix = row["prefix"].ToString();
                String name = row["name"].ToString();
                if (!_column3Items.ContainsKey(id))
                {
                    _column3Items.Add(id, prefix + ":" + name);
                }
            }


            foreach (KeyValuePair<String, String> k in _column3Items)
            {
                String item = k.Value;
                if (item.Substring(0, 1) == ":") item = item.Substring(1); //若選項開頭為":",擷取第二字元到結尾
                Column3.Items.Add(item); //建立Column3的選單
            }

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (tryConvert(comboBoxEx2.Text))
            {
                try
                {
                    SaveMappingRecord();
                    ReadMappingData();
                    DataSetting();
                }
                catch
                {
                    MessageBox.Show("網路或資料庫異常,請稍後再試...");
                    this.buttonX1.Enabled = true;
                    this.linkLabel1.Enabled = true;
                    this.dataGridViewX1.Enabled = true;
                    this.comboBoxEx1.Enabled = true;
                    this.comboBoxEx2.Enabled = true;
                }

            }

        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void ReadMappingData() //讀取DataGridView資料
        {
            SetMappingDataKey(); //初始化_mappingData
            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if (r.Cells[0].Value != null && r.Cells[1].Value != null)  //欄位有空值跳下一行
                {
                    String id = "";
                    foreach (KeyValuePair<String, String> k in _column3Items) //尋找選項的TagID
                    {
                        String item = r.Cells[1].Value.ToString();
                        if (!item.Contains(":")) //若選項無":"字串代表建立時prefix為空白,查詢時需補上":"
                        {
                            item = ":" + item;
                        }
                        if (item == k.Value)
                        {
                            id = k.Key;
                        }
                    }

                    if (id != "") //找不到對應ID不執行
                    {
                        if (!_mappingData.ContainsKey(r.Cells[0].Value.ToString())) //建立目標對應ID的字典
                        {
                            _mappingData.Add(r.Cells[0].Value.ToString(), new List<string>());
                        }
                        _mappingData[r.Cells[0].Value.ToString()].Add(id); //收集Mapping的TagId
                    }
                }
            }

            foreach (KeyValuePair<String, List<String>> k in _mappingData) //刪去value中的重複ID
            {
                for (int i = 0; i < k.Value.Count; i++)
                {
                    String s = k.Value[i];
                    int count = 0; //重複次數
                    for (int j = 0; j < k.Value.Count; j++)
                    {
                        if (s == k.Value[j]) //發現相同ID
                        {
                            count++;
                            if (count > 1) //只允許大於1
                            {
                                k.Value.Remove(s);
                                j--;
                                count--;
                            }
                        }
                    }
                }
            }
        }

        void DataSetting()
        {
            dept = this.comboBoxEx1.Text;
            _SchoolYear = comboBoxEx2.Text;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在產生新生入學方式統計表...");
            this.buttonX1.Enabled = false;
            this.linkLabel1.Enabled = false;
            this.dataGridViewX1.Enabled = false;
            this.comboBoxEx1.Enabled = false;
            this.comboBoxEx2.Enabled = false;
            _BGWClassStudentAbsenceDetail = new BackgroundWorker();
            _BGWClassStudentAbsenceDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentAbsenceDetail_DoWork);
            _BGWClassStudentAbsenceDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWClassStudentAbsenceDetail_Completed);
            _BGWClassStudentAbsenceDetail.RunWorkerAsync();
        }

        private void _BGWClassStudentAbsenceDetail_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            this.buttonX1.Enabled = true;
            this.linkLabel1.Enabled = true;
            this.dataGridViewX1.Enabled = true;
            this.comboBoxEx1.Enabled = true;
            this.comboBoxEx2.Enabled = true;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 新生入學方式統計表 已完成");

            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "新生入學方式統計表.xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _wk.Save(sd.FileName);
                    if (filter.error_list.Count > 0)
                    {
                        MessageBox.Show("發現" + filter.error_list.Count + "筆異常資料未列入統計\r\n詳細資料請確認報表中的[異常資料表]");
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

        private void _BGWClassStudentAbsenceDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            Dictionary<String, myStudent> myDic = new Dictionary<string, myStudent>();
            List<myStudent> mylist = new List<myStudent>();
            QueryHelper _Q = new QueryHelper();

            //SQL查詢要求的年級資料
            DataTable dt = _Q.Select("select student.id,student.name,student.gender,student.ref_class_id,student.status,class.class_name,class.grade_year,dept.name as dept_name,tag_student.ref_tag_id from student left join class on student.ref_class_id=class.id left join dept on class.ref_dept_id=dept.id left join tag_student on student.id= tag_student.ref_student_id where student.status in ('1','4','16') and class.grade_year='1'");

            //建立myStuden物件放至List中
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String name = row["name"].ToString();
                String gender = row["gender"].ToString();
                String ref_class_id = row["ref_class_id"].ToString();
                String class_name = row["class_name"].ToString();
                String grade_year = row["grade_year"].ToString();
                String dept_name = row["dept_name"].ToString();
                String ref_tag_id = row["ref_tag_id"].ToString();
                if (!myDic.ContainsKey(id)) //ID當key,不存在就建立
                {
                    myDic.Add(id, new myStudent(id, name, gender, ref_class_id, class_name, grade_year, dept_name, new List<string>()));
                }
                myDic[id].Tag.Add(ref_tag_id);
            }

            //取得新生異動資料清單
            List<K12.Data.UpdateRecordRecord> records = K12.Data.UpdateRecord.SelectByStudentIDs(myDic.Keys.ToList());
            foreach (KeyValuePair<String, myStudent> kvp in myDic)
            {
                if (CheckStudentStatus(records, kvp.Key)) //檢查新生異動資料是否符合
                {
                    mylist.Add(kvp.Value);  //符合者加入mylist清單
                }
            }

            filter = new Filter(mylist, dept);
            Export();
        }

        //確認學生為一般新生,排除重讀生等其他狀態
        public bool CheckStudentStatus(List<K12.Data.UpdateRecordRecord> records, String id)
        {
            foreach (K12.Data.UpdateRecordRecord record in records)
            {
                if (record.StudentID == id)  //找到符合的ID開始後續比對
                {
                    if (record.SchoolYear.ToString() == _SchoolYear) //確認學年度為當前學年度
                    {
                        if (Convert.ToInt16(record.UpdateCode) < 100) //異動代碼小於100
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        //輸出至Excel
        public void Export()
        {
            _wk = new Workbook();
            Worksheet ws;
            Cells cs;
            int index, row;
            _wk.Worksheets.Add();
            ////該年級學生總表
            //_wk.Worksheets.Add();
            //ws = _wk.Worksheets[2];
            //ws.Name = "All_List";
            //cs = ws.Cells;
            //cs["A1"].PutValue("ID");
            //cs["B1"].PutValue("Name");
            //cs["C1"].PutValue("Gender");
            //cs["D1"].PutValue("Ref_Class_Id");
            //cs["E1"].PutValue("Class_Name");
            //cs["F1"].PutValue("Grade_Year");
            //cs["G1"].PutValue("Dept_Name");
            //cs["H1"].PutValue("ref_tag_id");
            //index = 1;
            //foreach (myStudent s in filter.clean_list)
            //{
            //    cs[index, 0].PutValue(s.Id);
            //    cs[index, 1].PutValue(s.Name);
            //    cs[index, 2].PutValue(s.Gender);
            //    cs[index, 3].PutValue(s.Ref_class_id);
            //    cs[index, 4].PutValue(s.Class_name);
            //    cs[index, 5].PutValue(s.Grade_year);
            //    cs[index, 6].PutValue(s.Dept_name);
            //    String column7 = "";
            //    foreach (String l in s.Tag)
            //    {
            //        column7 += l + ",";
            //    }
            //    cs[index, 7].PutValue(column7);
            //    index++;
            //}

            //該科別異常的學生資料表


            ws = _wk.Worksheets[1];
            ws.Name = "異常資料表";
            cs = ws.Cells;
            cs["A1"].PutValue("系統編號");
            cs["B1"].PutValue("姓名");
            cs["C1"].PutValue("性別");
            //cs["D1"].PutValue("Ref_Class_Id");
            cs["D1"].PutValue("班級名稱");
            cs["E1"].PutValue("年級");
            cs["F1"].PutValue("科別名稱");
            //cs["H1"].PutValue("ref_tag_id");
            index = 1;
            foreach (myStudent s in filter.error_list)
            {
                cs[index, 0].PutValue(s.Id);
                cs[index, 1].PutValue(s.Name);
                cs[index, 2].PutValue(s.Gender);
                //cs[index, 3].PutValue(s.Ref_class_id);
                cs[index, 3].PutValue(s.Class_name);
                cs[index, 4].PutValue(s.Grade_year);
                cs[index, 5].PutValue(s.Dept_name);
                //String column7 = "";
                //foreach (String l in s.Tag)
                //{
                //    column7 += l + ",";
                //}
                //cs[index, 7].PutValue(column7);
                index++;
            }

            //新生入學方式統計表
            Workbook wk2 = new Workbook();
            wk2.Open(new MemoryStream(Properties.Resources.template)); //開啟範本文件

            _wk.Worksheets[0].Copy(wk2.Worksheets[0]); //複製範本文件
            ws = _wk.Worksheets[0];
            ws.Name = "新生入學方式統計表";
            cs = ws.Cells;

            index = 10;

            List<myStudent> summary = new List<myStudent>(); //建立summary清單收集dic_byDept的展開學生物件
            foreach (KeyValuePair<String, List<myStudent>> k in filter.dic_byDept)
            {
                //Table1 Left
                cs[index, 1].PutValue(filter.getDeptCode(k.Key)); //科別代碼
                cs[index, 2].PutValue(k.Key); //科別名稱
                cs[index, 6].PutValue(filter.getClassCount(k.Value)); //實際招生班數
                cs[index, 7].PutValue(k.Value.Count); //學生總計數
                cs[index, 8].PutValue(filter.getGenderCount(k.Value, "1")); //男生總數
                cs[index, 9].PutValue(filter.getGenderCount(k.Value, "0")); //女生總數

                foreach (myStudent s in k.Value)
                {
                    summary.Add(s); //展開dic_byDept,收集內容的myStudent物件
                }


                //Table1 Right
                row = 10;
                foreach (KeyValuePair<String, List<String>> map in _mappingData) //Form2傳入的Mapping資料
                {
                    if (map.Value.Count > 0)
                    {
                        List<myStudent> list = new List<myStudent>();
                        list = filter.getListByTagId(map.Value, k.Value); //list收集符合的TagId學生物件
                        cs[index, row].PutValue(list.Count); //列出符合的TagId學生物件總數
                    }
                    row++; //換欄

                }
                index++; //每做完一次k.value即換行
            }

            //Table2 Left
            Dictionary<String, List<String>> table2Left = new Dictionary<string, List<String>>();

            foreach (KeyValuePair<String, List<String>> map in _mappingData)
            {
                String[] key = map.Key.Split(':');
                if (!table2Left.ContainsKey(key[1]))
                {
                    table2Left.Add(key[1], new List<String>());
                }
                foreach (String s in map.Value)
                {
                    if (map.Key.Split(':')[1] == key[1])
                    {
                        table2Left[key[1]].Add(s);
                    }
                }
            }

            //收集原住民生
            foreach (KeyValuePair<String, List<String>> k in table2Left)
            {
                if (k.Key == "原住民生")
                {
                    AboList = k.Value; //收入TagID
                }
            }

            index = 32;
            foreach (KeyValuePair<String, List<String>> k in table2Left)
            {
                List<myStudent> list = new List<myStudent>();
                list = filter.getListByTagId(k.Value, summary);

                cs[index, 4].PutValue(list.Count);
                cs[index, 6].PutValue(filter.getGenderCount(list, "1"));
                cs[index, 8].PutValue(filter.getGenderCount(list, "0"));
                index++;
            }

            //Table2 Right
            index = 32;
            row = 10;
            foreach (KeyValuePair<String, List<String>> map in _mappingData)
            {
                if (index > 35) { index = 32; row += 2; } //換行換欄
                if (map.Value.Count > 0)
                {
                    List<myStudent> list = new List<myStudent>();
                    list = filter.getListByTagId(map.Value, summary);

                    cs[index, row].PutValue(filter.getGenderCount(list, "1"));
                    cs[index, row + 1].PutValue(filter.getGenderCount(list, "0"));
                }
                index++;


            }

            //Table3 Left
            List<myStudent> collect__LastGradeT = new List<myStudent>();  //應屆的收集清單
            List<myStudent> collect__LastGradeF = new List<myStudent>();  //非應屆的收集清單
            List<String> collect_List = new List<string>(); //收集學生ID的清單

            foreach (myStudent student in summary) //收集summary所有學生ID
            {
                collect_List.Add(student.Id);
            }
            //傳入學生ID清單供查詢
            List<SHSchool.Data.SHBeforeEnrollmentRecord> recl = SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
            foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
            {
                foreach (myStudent student in summary)
                {
                    if (rec.RefStudentID == student.Id) //找到對應ID後,判斷前級畢業年度
                    {
                        String last_grade_year = rec.GraduateSchoolYear;
                        if (last_grade_year == "") last_grade_year = "0"; //空值填方便後續計算
                        int year = Convert.ToInt16(last_grade_year) + 1912; //學年度+1912若等於現在年份則判斷為應屆生
                        if (year.ToString() == DateTime.Now.Year.ToString())
                        {
                            collect__LastGradeT.Add(student); //收入應屆清單
                        }
                        else
                        {
                            collect__LastGradeF.Add(student); //收入非應屆清單
                        }
                    }
                }
            }

            cs[36, 4].PutValue(collect__LastGradeT.Count); //應屆畢業總數
            cs[37, 4].PutValue(collect__LastGradeF.Count); //非應屆畢業總數
            cs[36, 6].PutValue(filter.getGenderCount(collect__LastGradeT, "1")); //應屆畢業男生總數
            cs[36, 8].PutValue(filter.getGenderCount(collect__LastGradeT, "0")); //應屆畢業女生總數
            cs[37, 6].PutValue(filter.getGenderCount(collect__LastGradeF, "1")); //非應屆畢業男生總數
            cs[37, 8].PutValue(filter.getGenderCount(collect__LastGradeF, "0")); //非應屆畢業女生總數

            //Table3 Right
            Dictionary<String, List<String>> ndic = new Dictionary<string, List<string>>(); //為綜合入學方式,建立字典
            foreach (KeyValuePair<String, List<String>> map in _mappingData)
            {
                String key = map.Key.Substring(0, 2); //建立key為前面兩個字串:甄選,申請,登記,直升,免試,其他
                if (!ndic.ContainsKey(key))
                {
                    ndic.Add(key, new List<string>()); //key不存在即建立
                }
                foreach (String s in map.Value)
                {
                    if (map.Key.Contains(key)) //針對符合的key做TagID的收集
                    {
                        ndic[key].Add(s);
                    }
                }
            }

            index = 36;
            row = 10;
            foreach (KeyValuePair<String, List<String>> nmap in ndic)
            {
                if (index > 36) { index = 36; row += 2; } //換行換欄
                if (nmap.Value.Count == 0)  //遇到空值index++並繼續迴圈
                {
                    index++;
                    continue;
                }
                List<myStudent> list = new List<myStudent>();
                list = filter.getListByTagId(nmap.Value, summary); //收集符合TagID的學生物件
                collect_List = new List<string>(); //清空之前的清單
                collect__LastGradeT = new List<myStudent>(); //清空之前的清單
                collect__LastGradeF = new List<myStudent>(); //清空之前的清單
                foreach (myStudent student in list)
                {
                    collect_List.Add(student.Id); //收集學生ID
                }

                recl = SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
                foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
                {
                    foreach (myStudent student in list)
                    {
                        if (rec.RefStudentID == student.Id)
                        {
                            String last_grade_year = rec.GraduateSchoolYear;
                            if (last_grade_year == "") last_grade_year = "0";
                            int year = Convert.ToInt16(last_grade_year) + 1912;
                            if (year.ToString() == DateTime.Now.Year.ToString())
                            {
                                collect__LastGradeT.Add(student); //收入應屆清單
                            }
                            else
                            {
                                collect__LastGradeF.Add(student); //收入非應屆清單
                            }
                        }
                    }
                }
                cs[index, row].PutValue(filter.getGenderCount(collect__LastGradeT, "1")); //應屆男生數
                cs[index, row + 1].PutValue(filter.getGenderCount(collect__LastGradeT, "0")); //應屆女生數
                cs[index + 1, row].PutValue(filter.getGenderCount(collect__LastGradeF, "1")); //非應屆男生數
                cs[index + 1, row + 1].PutValue(filter.getGenderCount(collect__LastGradeF, "0")); //非應屆女生數
                index++; //換行
            }

            //Table3 End
            collect_List = new List<string>(); //清空之前的清單
            collect__LastGradeT = new List<myStudent>(); //清空之前的清單
            collect__LastGradeF = new List<myStudent>(); //清空之前的清單
            List<myStudent> AboStudent = filter.getListByTagId(AboList, summary);
            foreach (myStudent student in AboStudent)
            {
                collect_List.Add(student.Id);
            }
            recl = SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
            foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
            {
                foreach (myStudent student in AboStudent)
                {
                    if (rec.RefStudentID == student.Id)
                    {
                        String last_grade_year = rec.GraduateSchoolYear;
                        if (last_grade_year == "") last_grade_year = "0";
                        int year = Convert.ToInt16(last_grade_year) + 1912;
                        if (year.ToString() == DateTime.Now.Year.ToString())
                        {
                            collect__LastGradeT.Add(student); //收入應屆清單
                        }
                        else
                        {
                            collect__LastGradeF.Add(student); //收入非應屆清單
                        }
                    }
                }
            }

            cs[36, 22].PutValue(filter.getGenderCount(collect__LastGradeT, "1")); //應屆原住民男生數
            cs[36, 23].PutValue(filter.getGenderCount(collect__LastGradeT, "0")); //應屆原住民女生數
            cs[37, 22].PutValue(filter.getGenderCount(collect__LastGradeF, "1")); //非應屆原住民男生數
            cs[37, 23].PutValue(filter.getGenderCount(collect__LastGradeF, "0")); //非應屆原住民女生數
            cs["U5"].PutValue(_SchoolYear);
        }

        public void SaveMappingRecord() //儲存上次Mapping紀錄
        {
            AccessHelper _A = new AccessHelper();
            List<myTableUDT> UDTlist = _A.Select<myTableUDT>();
            _A.DeletedValues(UDTlist); //清除UDT資料

            UDTlist = new List<myTableUDT>(); //清空UDTlist
            foreach (DataGridViewRow row in dataGridViewX1.Rows) //取得DataDataGridViewRow資料
            {
                if (row.Cells[0].Value == null) //遇到空白的Target即跳到下個loop
                {
                    continue;
                }

                String target = row.Cells[0].Value.ToString();
                String source = "";
                if (row.Cells[1].Value != null) { source = row.Cells[1].Value.ToString(); }

                myTableUDT obj = new myTableUDT();
                obj.Target = target;
                obj.Source = source;
                UDTlist.Add(obj);
            }

            _A.InsertValues(UDTlist); //回存到UDT
        }

        public void LoadLastRecord() //讀取上次Mapping設定
        {
            AccessHelper _A = new AccessHelper();
            List<myTableUDT> UDTlist = _A.Select<myTableUDT>(); //檢查UDT並回傳資料
            DataGridViewRow row;
            if (UDTlist.Count > 0) //UDT內有設定才做讀取
            {
                for (int i = 0; i < UDTlist.Count; i++)
                {
                    row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = UDTlist[i].Target;
                    row.Cells[1].Value = UDTlist[i].Source;
                    dataGridViewX1.Rows.Add(row);
                }
            }
            else
            {
                //UDT無資料則提供預設標記
                for (int i = 0; i < Column2.Items.Count; i++)
                {
                    row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = Column2.Items[i];
                    dataGridViewX1.Rows.Add(row);
                }
            }

        }

        //初始化MappingData資料
        public void SetMappingDataKey()
        {
            _mappingData = new Dictionary<string, List<string>>();
            foreach (String s in Column2.Items)
            {
                _mappingData.Add(s, new List<string>());
            }
        }

        private void SchoolYearItem() //建立comboBoxEx2的下拉清單
        {
            int school_year = Convert.ToInt16(K12.Data.School.DefaultSchoolYear);
            for (int i = -3; i < 4; i++)
            {
                comboBoxEx2.Items.Add(school_year + i);
            }
            comboBoxEx2.Text = K12.Data.School.DefaultSchoolYear;
        }

        private bool tryConvert(String str) //避免comboBoxEx2被輸入非數字字串
        {
            try
            {
                Convert.ToInt16(str);
                return true;
            }
            catch
            {
                MessageBox.Show("學年度請確認輸入正確數值");
                return false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                AccessHelper _A = new AccessHelper();
                List<myTableUDT> UDTlist = _A.Select<myTableUDT>();
                _A.DeletedValues(UDTlist); //清除UDT資料
                dataGridViewX1.Rows.Clear();  //清除datagridview資料
                LoadLastRecord(); //再次讀入Mapping設定
            }
            catch
            {
                MessageBox.Show("網路或資料庫異常,請稍後再試...");
            }
            
        }
    }

}




