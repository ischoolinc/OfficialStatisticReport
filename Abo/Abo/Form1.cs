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

namespace Abo
{
    public partial class Form1 : BaseForm
    {
        Dictionary<String, String> _column2Items,_Group; //全部類別對照表,key=TagId,value=prefix+":"+name
        Dictionary<String, List<String>> _mappingData;//mapping資料
        List<String> _TagIDList; //被選取到的TagID總表,排除學生清單用
        List<myStudent> _CleanList, _ErrorList, 普通科, 綜合高中科, 職業科;
        List<GraduateStudentObj> _GraduateStudentList,_GCleanList, _GErrorList;
        private BackgroundWorker _BGWClassStudentAbsenceDetail; //背景模式
        Workbook _wk;
        String _SchoolYear;

        public Form1()
        {
            InitializeComponent();
            Column1Prepare();
            Column2Prepare();
            LoadLastRecord();
        }

        private void Column1Prepare()
        {
            _Group = new Dictionary<string, string>();
            _Group.Add("21","阿美族");
            _Group.Add("22","泰雅族");
            _Group.Add("23","排灣族");
            _Group.Add("24","布農族");
            _Group.Add("25","卑南族");
            _Group.Add("26","鄒(曹)族");
            _Group.Add("27","魯凱族");
            _Group.Add("28","賽夏族");
            _Group.Add("29","雅美族");
            _Group.Add("2A","卲族");
            _Group.Add("2B","噶嗎蘭族");
            _Group.Add("2C","太魯閣族");
            _Group.Add("2D","撒奇萊雅族");
            _Group.Add("2E","賽德克族");
            _Group.Add("20","其他");

            foreach(KeyValuePair<String,String> k in _Group)
            {
                    Column1.Items.Add(k.Value);
            }
        }

        private void Column2Prepare()
        {
            _column2Items = new Dictionary<String, String>();
            QueryHelper _Q = new QueryHelper();

            DataTable dt = _Q.Select("select * from tag where category='Student' order by prefix,name");
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String prefix = row["prefix"].ToString();
                String name = row["name"].ToString();
                if (!_column2Items.ContainsKey(id))
                {
                    _column2Items.Add(id, prefix + ":" + name);
                }
            }

            foreach (KeyValuePair<String, String> k in _column2Items)
            {
                String item = k.Value;
                if (item.Substring(0, 1) == ":") item = item.Substring(1); //若選項開頭為":",擷取第二字元到結尾
                Column2.Items.Add(item); //建立Column2的選單
            }
        }

        public void LoadLastRecord() //讀取上次Mapping設定
        {
            AccessHelper _A = new AccessHelper();
            List<AboTableUDT> UDTlist = _A.Select<AboTableUDT>(); //檢查UDT並回傳資料
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

        private void buttonX1_Click(object sender, EventArgs e)
        {
            try
            {
                SaveMappingRecord();
                ReadMappingData();
                SetTagIDList();
                DataSetting();
            }
            catch
            {
                MessageBox.Show("網路或資料庫異常,請稍後再試...");
                this.buttonX1.Enabled = true;
                this.linkLabel1.Enabled = true;
                this.dataGridViewX1.Enabled = true;
            }
        }

        public void SaveMappingRecord() //儲存上次Mapping紀錄
        {
            AccessHelper _A = new AccessHelper();
            List<AboTableUDT> UDTlist = _A.Select<AboTableUDT>();
            _A.DeletedValues(UDTlist); //清除UDT資料

            UDTlist = new List<AboTableUDT>(); //清空UDTlist
            foreach (DataGridViewRow row in dataGridViewX1.Rows) //取得DataDataGridViewRow資料
            {
                if (row.Cells[0].Value.ToString() == "") //遇到空白的Target即跳到下個loop
                {
                    continue;
                }

                String target = row.Cells[0].Value.ToString();
                String source = "";
                if (row.Cells[1].Value != null) { source = row.Cells[1].Value.ToString(); }

                AboTableUDT obj = new AboTableUDT();
                obj.Target = target;
                obj.Source = source;
                UDTlist.Add(obj);
            }

            _A.InsertValues(UDTlist); //回存到UDT
        }

        void ReadMappingData() //讀取DataGridView資料
        {
            _mappingData = new Dictionary<string, List<string>>();

            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if (r.Cells[0].Value != null && r.Cells[1].Value != null)  //欄位有空值跳下一行
                {
                    String id = "";
                    foreach (KeyValuePair<String, String> k in _column2Items) //尋找選項的TagID
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在產生原住民學生數及畢業生統計表...");
            this.buttonX1.Enabled = false;
            this.linkLabel1.Enabled = false;
            this.dataGridViewX1.Enabled = false;
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 原住民學生數及畢業生統計表 已完成");

            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "原住民學生數及畢業生統計表.xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _wk.Save(sd.FileName);
                    if ((_ErrorList.Count + _GErrorList.Count)> 0)
                    {
                        MessageBox.Show("發現" + (_ErrorList.Count + _GErrorList.Count) + "筆異常資料未列入統計\r\n詳細資料請確認報表中的[異常資料表]");
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
            _SchoolYear = K12.Data.School.DefaultSchoolYear;
            _GraduateStudentList = getGraduateStudent();

            Dictionary<String, myStudent> myDic = new Dictionary<string, myStudent>();
            List<myStudent> mylist = new List<myStudent>();
            QueryHelper _Q = new QueryHelper();

            //SQL查詢要求的年級資料
            DataTable dt = _Q.Select("select student.id,student.name,student.student_number,student.gender,student.ref_class_id,student.status,class.class_name,class.grade_year,dept.name as dept_name,tag_student.ref_tag_id from student left join class on student.ref_class_id=class.id left join dept on class.ref_dept_id=dept.id left join tag_student on student.id= tag_student.ref_student_id where student.status in ('1','2')");

            //建立myStuden物件放至List中
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String name = row["name"].ToString();
                String student_number = row["student_number"].ToString();
                String gender = row["gender"].ToString();
                String ref_class_id = row["ref_class_id"].ToString();
                String class_name = row["class_name"].ToString();
                String grade_year = row["grade_year"].ToString();
                String dept_name = row["dept_name"].ToString();
                String ref_tag_id = row["ref_tag_id"].ToString();
                String status = row["status"].ToString();
                if (!myDic.ContainsKey(id)) //ID當key,不存在就建立
                {
                    myDic.Add(id, new myStudent(id, name, student_number, gender, ref_class_id, class_name, grade_year, dept_name, status, new List<string>()));
                }
                myDic[id].Tag.Add(ref_tag_id);
            }

            //Dic to List
            mylist = myDic.Values.ToList();
            
            Cleaner(mylist);
            Export();
        }

        private void Cleaner(List<myStudent> list)
        {
            _CleanList = new List<myStudent>();
            _ErrorList = new List<myStudent>();

            foreach (myStudent s in list)
            {
                if (!TagIDExistence(s)) continue; //無資料不需要加入就換下一個學生

                //按照資料的正確性分別加入Error或CleanList
                if (s.Id == "" || s.Name == "" || (s.Gender != "0" && s.Gender != "1") || s.Ref_class_id == "" || s.Class_name == "" || s.Grade_year == "" || s.Dept_name == "")
                {
                    _ErrorList.Add(s);
                }
                else
                {
                    _CleanList.Add(s);
                }
            }

            普通科 = getStudentListByDept("普通科");
            綜合高中科 = getStudentListByDept("綜合高中科");
            職業科 = getStudentListByDept("職業科");
        }

        private void Export()
        {
            _wk = new Workbook();
            Workbook wk2 = new Workbook();
            wk2.Open(new MemoryStream(Properties.Resources.aboriginal)); //開啟範本文件
            _wk.Worksheets[0].Copy(wk2.Worksheets[0]); //複製範本文件
            Worksheet ws = _wk.Worksheets[0];
            ws.Name = "普通科";
            Cells cs = ws.Cells;
            int index = 9;
            Dictionary<String, List<myStudent>> dica = getStudentDic(普通科);
            foreach (KeyValuePair<String, List<myStudent>> k in dica)
            {
                cs[index, 0].PutValue(getGroupCode(k.Key));
                cs[index, 1].PutValue(k.Key);
                cs[index, 2].PutValue(getStudentCount(k.Value));
                cs[index, 3].PutValue(getStudentCount(k.Value, "0", "1"));
                cs[index, 4].PutValue(getStudentCount(k.Value, "0", "0"));
                cs[index, 5].PutValue(getStudentCount(k.Value, "1", "1"));
                cs[index, 6].PutValue(getStudentCount(k.Value, "1", "0"));
                cs[index, 7].PutValue(getStudentCount(k.Value, "2", "1"));
                cs[index, 8].PutValue(getStudentCount(k.Value, "2", "0"));
                cs[index, 9].PutValue(getStudentCount(k.Value, "3", "1"));
                cs[index, 10].PutValue(getStudentCount(k.Value, "3", "0"));
                cs[index, 11].PutValue(getStudentCount(k.Value, "4", "1"));
                cs[index, 12].PutValue(getStudentCount(k.Value, "4", "0"));
                cs[index, 13].PutValue(getStudentCount(k.Value, "5", "1"));
                cs[index, 14].PutValue(getStudentCount(k.Value, "5", "0"));

                List<GraduateStudentObj> list = getGraduateStudentList(k.Key, "普通科");
                cs[index, 15].PutValue(list.Count);
                cs[index, 16].PutValue(getGraduateStudentCount(list, "1"));
                cs[index, 17].PutValue(getGraduateStudentCount(list, "0"));
                index++;
            }

            _wk.Worksheets.Add();
            _wk.Worksheets[1].Copy(wk2.Worksheets[0]); //複製範本文件
            ws = _wk.Worksheets[1];
            ws.Name = "綜合高中科";
            cs = ws.Cells;
            index = 9;
            Dictionary<String, List<myStudent>> dicb = getStudentDic(綜合高中科);
            foreach (KeyValuePair<String, List<myStudent>> k in dicb)
            {
                cs[index, 0].PutValue(getGroupCode(k.Key));
                cs[index, 1].PutValue(k.Key);
                cs[index, 2].PutValue(getStudentCount(k.Value));
                cs[index, 3].PutValue(getStudentCount(k.Value, "0", "1"));
                cs[index, 4].PutValue(getStudentCount(k.Value, "0", "0"));
                cs[index, 5].PutValue(getStudentCount(k.Value, "1", "1"));
                cs[index, 6].PutValue(getStudentCount(k.Value, "1", "0"));
                cs[index, 7].PutValue(getStudentCount(k.Value, "2", "1"));
                cs[index, 8].PutValue(getStudentCount(k.Value, "2", "0"));
                cs[index, 9].PutValue(getStudentCount(k.Value, "3", "1"));
                cs[index, 10].PutValue(getStudentCount(k.Value, "3", "0"));
                cs[index, 11].PutValue(getStudentCount(k.Value, "4", "1"));
                cs[index, 12].PutValue(getStudentCount(k.Value, "4", "0"));
                cs[index, 13].PutValue(getStudentCount(k.Value, "5", "1"));
                cs[index, 14].PutValue(getStudentCount(k.Value, "5", "0"));

                List<GraduateStudentObj> list = getGraduateStudentList(k.Key, "綜合高中科");
                cs[index, 15].PutValue(list.Count);
                cs[index, 16].PutValue(getGraduateStudentCount(list, "1"));
                cs[index, 17].PutValue(getGraduateStudentCount(list, "0"));
                index++;
            }

            _wk.Worksheets.Add();
            _wk.Worksheets[2].Copy(wk2.Worksheets[0]); //複製範本文件
            ws = _wk.Worksheets[2];
            ws.Name = "職業科";
            cs = ws.Cells;
            index = 9;
            Dictionary<String, List<myStudent>> dicc = getStudentDic(職業科);
            foreach (KeyValuePair<String, List<myStudent>> k in dicc)
            {
                cs[index, 0].PutValue(getGroupCode(k.Key));
                cs[index, 1].PutValue(k.Key);
                cs[index, 2].PutValue(getStudentCount(k.Value));
                cs[index, 3].PutValue(getStudentCount(k.Value, "0", "1"));
                cs[index, 4].PutValue(getStudentCount(k.Value, "0", "0"));
                cs[index, 5].PutValue(getStudentCount(k.Value, "1", "1"));
                cs[index, 6].PutValue(getStudentCount(k.Value, "1", "0"));
                cs[index, 7].PutValue(getStudentCount(k.Value, "2", "1"));
                cs[index, 8].PutValue(getStudentCount(k.Value, "2", "0"));
                cs[index, 9].PutValue(getStudentCount(k.Value, "3", "1"));
                cs[index, 10].PutValue(getStudentCount(k.Value, "3", "0"));
                cs[index, 11].PutValue(getStudentCount(k.Value, "4", "1"));
                cs[index, 12].PutValue(getStudentCount(k.Value, "4", "0"));
                cs[index, 13].PutValue(getStudentCount(k.Value, "5", "1"));
                cs[index, 14].PutValue(getStudentCount(k.Value, "5", "0"));

                List<GraduateStudentObj> list = getGraduateStudentList(k.Key, "職業科");
                cs[index, 15].PutValue(list.Count);
                cs[index, 16].PutValue(getGraduateStudentCount(list, "1"));
                cs[index, 17].PutValue(getGraduateStudentCount(list, "0"));
                index++;
            }

            _wk.Worksheets.Add();
            ws = _wk.Worksheets[3];
            ws.Name = "異常資料表";
            cs = ws.Cells;
            cs["A1"].PutValue("學號");
            cs["B1"].PutValue("姓名");
            cs["C1"].PutValue("性別");
            cs["D1"].PutValue("班級名稱");
            cs["E1"].PutValue("年級");
            cs["F1"].PutValue("科別名稱");
            cs["G1"].PutValue("狀態");
            index = 1;
            foreach (myStudent s in _ErrorList)
            {
                cs[index, 0].PutValue(s.Student_number);
                cs[index, 1].PutValue(s.Name);
                cs[index, 2].PutValue(s.Gender);
                cs[index, 3].PutValue(s.Class_name);
                cs[index, 4].PutValue(s.Grade_year);
                cs[index, 5].PutValue(s.Dept_name);
                cs[index, 6].PutValue(s.Status == "1"?"一般生":"延修生");
                index++;
            }

            foreach(GraduateStudentObj obj in _GErrorList)
            {
                cs[index, 0].PutValue(obj.Student_number);
                cs[index, 1].PutValue(obj.Name);
                cs[index, 2].PutValue(obj.Gender);
                cs[index, 5].PutValue(obj.Dept);
                cs[index, 6].PutValue("畢業或離校生");
            }
        }

        private List<myStudent> getStudentListByDept(String dept)
        {
            List<myStudent> list = new List<myStudent>();

            switch (dept)
            {
                case "職業科":
                    foreach (myStudent s in _CleanList)
                    {
                        if (!s.Dept_name.Contains("普通科") && !s.Dept_name.Contains("綜合高中科"))
                        {
                            list.Add(s);
                        }
                    }
                    break;

                default:
                    foreach (myStudent s in _CleanList)
                    {
                        if (s.Dept_name.Contains(dept))
                        {
                            list.Add(s);
                        }
                    }
                    break;
            }
            return list;
        }

        private Dictionary<String, List<myStudent>> getStudentDic(List<myStudent> list)
        {
            Dictionary<String, List<myStudent>> dic = new Dictionary<string, List<myStudent>>();
            foreach (KeyValuePair<String, List<String>> k in _mappingData)
            {
                if (k.Value.Count == 0)
                {
                    continue;
                }

                if (!dic.ContainsKey(k.Key))
                {
                    dic.Add(k.Key, new List<myStudent>());
                }

                foreach (String tagid in k.Value)
                {
                    foreach (myStudent s in list)
                    {
                        foreach (String tag in s.Tag)
                        {
                            if (tag == tagid)
                            {
                                dic[k.Key].Add(s);
                                break;
                            }
                        }
                    }
                }
            }
            return dic;
        }

        //取得一般生及延修生的總數
        private int getStudentCount(List<myStudent> list)
        {
            int count = 0;
            foreach (myStudent s in list)
            {
                if (s.Status == "1" || s.Status == "2") 
                {
                    count++;
                }
            }
            return count++;
        }

        //取得指定年級性別的一般生總數
        private int getStudentCount(List<myStudent> list, String grade, String gender)
        {
            int count = 0;
            switch (grade)
            {
                case "0": //不指定年級
                    foreach (myStudent s in list)
                    {
                        if (s.Gender == gender & (s.Status == "1" || s.Status == "2"))
                        {
                            count++;
                        }
                    }
                    break;

                case "5": //延修生
                    String status = "2";
                    foreach (myStudent s in list)
                    {
                        if (s.Status == status && s.Gender == gender)
                        {
                            count++;
                        }
                    }
                    break;

                default:
                    foreach (myStudent s in list)
                    {
                        if (s.Grade_year == grade && s.Gender == gender && s.Status == "1")
                        {
                            count++;
                        }
                    }
                    break;
            }
            return count;
        }

        //取得上學年畢業生物件清單
        private List<GraduateStudentObj> getGraduateStudent()
        {
            _GCleanList = new List<GraduateStudentObj>();
            _GErrorList = new List<GraduateStudentObj>();
            int year = Convert.ToInt32(_SchoolYear) - 1; //當前系統學年度-1
            Dictionary<String, GraduateStudentObj> dic = new Dictionary<string, GraduateStudentObj>();
            FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();
            DataTable dt = _Q.Select("select update_record.ref_student_id,update_record.ss_name,student.student_number,update_record.ss_gender,update_record.ss_dept,tag_student.ref_tag_id from update_record left join tag_student on update_record.ref_student_id = tag_student.ref_student_id left join student on update_record.ref_student_id=student.id where update_code='501' and school_year='"+ year +"'");

            foreach (DataRow row in dt.Rows)
            {
                String id = row["ref_student_id"].ToString();
                String name = row["ss_name"].ToString();
                String student_number = row["student_number"].ToString();
                String gender = row["ss_gender"].ToString();
                String dept = row["ss_dept"].ToString();
                String tagid = row["ref_tag_id"].ToString();
                if(!dic.ContainsKey(id))
                {
                    dic.Add(id, new GraduateStudentObj(id, name, student_number,gender, dept, new List<String>()));
                }
                dic[id].TagID.Add(tagid);
            }

            //判斷性別欄位是否異常
            foreach(String id in dic.Keys)
            {
                if(dic[id].Gender != "1" && dic[id].Gender != "0")
                {
                    _GErrorList.Add(dic[id]);
                }
                else
                {
                    _GCleanList.Add(dic[id]);
                }
            }

            return _GCleanList;
        }

        //取得指定族別科別的上屆畢業生清單
        private List<GraduateStudentObj> getGraduateStudentList(String group_name,String dept)
        {
            List<GraduateStudentObj> list = new List<GraduateStudentObj>();
            
            switch(dept)
            {
                case "職業科":
                    foreach (KeyValuePair<String, List<String>> map in _mappingData)
                    {
                        if (map.Key == group_name) //搜尋指定族別的tagID
                        {
                            foreach (String map_id in map.Value)
                            {
                                foreach (GraduateStudentObj obj in _GraduateStudentList)
                                {
                                    if (!obj.Dept.Contains("普通科") && !obj.Dept.Contains("綜合高中科")) //從畢業生總清單搜尋指定的科別
                                    {
                                        foreach (String s in obj.TagID)
                                        {
                                            if (s == map_id) //判斷該學生的tagID是否符合
                                            {
                                                list.Add(obj);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    break;

                default:
                    foreach (KeyValuePair<String, List<String>> map in _mappingData)
                    {
                        if (map.Key == group_name) //搜尋指定族別的tagID
                        {
                            foreach (String map_id in map.Value)
                            {
                                foreach (GraduateStudentObj obj in _GraduateStudentList)
                                {
                                    if (obj.Dept.Contains(dept)) //從畢業生總清單搜尋指定的科別
                                    {
                                        foreach (String s in obj.TagID) 
                                        {
                                            if (s == map_id) //判斷該學生的tagID是否符合
                                            {
                                                list.Add(obj);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    break;
            }
            
            return list;
        }

        //取得清單中指定性別的畢業生總數
        private int getGraduateStudentCount(List<GraduateStudentObj> list,String gender)
        {
            int count = 0;
            foreach(GraduateStudentObj obj in list)
            {
                if(obj.Gender == gender)
                {
                    count++;
                }
            }
            return count;
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                AccessHelper _A = new AccessHelper();
                List<AboTableUDT> UDTlist = _A.Select<AboTableUDT>();
                _A.DeletedValues(UDTlist); //清除UDT資料
                dataGridViewX1.Rows.Clear();  //清除datagridview資料
                LoadLastRecord(); //再次讀入Mapping設定
            }
            catch
            {
                MessageBox.Show("網路或資料庫異常,請稍後再試...");
            }
        }

        //取得族別代號
        private String getGroupCode(String groupName)
        {
            String code = "";
            foreach(KeyValuePair<String,String> k in _Group)
            {
                if(k.Value == groupName)
                {
                    code = k.Key;
                }
            }
            return code;
        }

        //儲存被選取的TagID
        private void SetTagIDList()
        {
            _TagIDList = new List<string>();
            foreach(String key in _mappingData.Keys)
            {
                foreach(String tagid in _mappingData[key])
                {
                    _TagIDList.Add(tagid);
                }
            }
        }

        //檢查該學生的tagid是否在清單中
        private bool TagIDExistence(myStudent student)
        {
            bool addOrNot = false;  //是否加入清單的判斷值
            foreach (String tagid in student.Tag)
            {
                if (_TagIDList.Contains(tagid)) //查詢此學生的tagid是否有在_TagIDLIst中,若有代表需要這筆資料
                {
                    addOrNot = true;
                    break; //有找到即可跳離,無須再比對後續tagid
                }
            }
            return addOrNot;
        }
    }
}
