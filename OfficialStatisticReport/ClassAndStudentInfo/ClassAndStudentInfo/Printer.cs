using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using SHSchool.Data;
using Aspose.Cells;
using System.IO;
using FISCA.Data;

namespace ClassAndStudentInfo
{
    class Printer
    {
        String _SchoolYear; //當前學年度
        private List<StudentObj> _ErrorList, _CorrectList;
        private List<GraduateStudentObj> _GraduateStudentList;
        private List<ReStudentObj> _ReStudentList;
        private Dictionary<String, List<StudentObj>> DeptDic, 普通科, 綜合高中科, 職業科;
        Dictionary<String, String> Dept_ref; //科別代碼對照,key=code,value=name;
        private BackgroundWorker _BGWClassStudentAbsenceDetail; //背景模式
        Workbook _Wk;

        public void Start()
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在產生班級及學生概況統計表...");
            _BGWClassStudentAbsenceDetail = new BackgroundWorker();
            _BGWClassStudentAbsenceDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentAbsenceDetail_DoWork);
            _BGWClassStudentAbsenceDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWClassStudentAbsenceDetail_Completed);
            _BGWClassStudentAbsenceDetail.RunWorkerAsync();
        }

        private void _BGWClassStudentAbsenceDetail_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 班級及學生概況統計表 已完成");
            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "班級及學生概況統計表.xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _Wk.Save(sd.FileName);
                    if (_ErrorList.Count > 0)
                    {
                        MessageBox.Show("發現" + _ErrorList.Count + "筆異常資料未列入統計\r\n詳細資料請確認報表中的[異常資料表]");
                    }
                    System.Diagnostics.Process.Start(sd.FileName);

                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }
            }
        }


        private void _BGWClassStudentAbsenceDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            _SchoolYear = K12.Data.School.DefaultSchoolYear;
            _GraduateStudentList = getGraduateStudent();
            _ReStudentList = getReStudent();
            QueryDeptCode(); //建立科別代號表

            FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();
            DataTable dt = _Q.Select("select student.id,student.name,student.ref_class_id,student.status,gender,class.grade_year,dept.name as dept_name from student join class on student.ref_class_id=class.id join dept on class.ref_dept_id=dept.id");
            List<StudentObj> StudentList = new List<StudentObj>();
            foreach (DataRow row in dt.Rows)
            {
                StudentObj obj = new StudentObj(row);
                StudentList.Add(obj);
            }

            //資料整理
            CleanList(StudentList);
            DeptDic = SortToDept(_CorrectList); //科別分類
            普通科 = getDicByDept(DeptDic, "普通科");
            綜合高中科 = getDicByDept(DeptDic, "綜合高中科");
            職業科 = getDicByDept(DeptDic, "職業科");

            //資料列印           
            Export();
        }

        //科別分類,key=科別名稱,value=學生物件清單
        private Dictionary<String, List<StudentObj>> SortToDept(List<StudentObj> StudentList)
        {
            Dictionary<String, List<StudentObj>> Dic = new Dictionary<string, List<StudentObj>>();
            foreach (StudentObj obj in StudentList)
            {
                if (!Dic.ContainsKey(obj.科別)) //建立科別
                {
                    Dic.Add(obj.科別, new List<StudentObj>());
                }
                Dic[obj.科別].Add(obj);
            }
            return Dic;
        }



        private void Export()
        {
            _Wk = new Workbook();
            _Wk.Worksheets.Add();
            _Wk.Worksheets.Add();
            _Wk.Worksheets.Add();
            _Wk.Worksheets[0].Copy(Template()); //複製範本
            _Wk.Worksheets[1].Copy(Template()); //複製範本
            _Wk.Worksheets[2].Copy(Template()); //複製範本
            Worksheet ws = _Wk.Worksheets[3];
            Cells cs;

            ws.Name = "異常資料表";
            cs = ws.Cells;
            //異常資料表
            cs[0, 0].PutValue("系統編號");
            cs[0, 1].PutValue("姓名");
            cs[0, 2].PutValue("班級編號");
            cs[0, 3].PutValue("性別");
            cs[0, 4].PutValue("年級");
            cs[0, 5].PutValue("科別");
            int index = 1;

            foreach (StudentObj obj in _ErrorList)
            {
                cs[index, 0].PutValue(obj.學生ID);
                cs[index, 1].PutValue(obj.學生名字);
                cs[index, 2].PutValue(obj.班級ID);
                cs[index, 3].PutValue(obj.性別);
                cs[index, 4].PutValue(obj.年級);
                cs[index, 5].PutValue(obj.科別);
                index++;
            }

            //普通科
            ws = _Wk.Worksheets[0];
            ws.Name = "普通科";
            cs = ws.Cells;
            index = 9;
            foreach (KeyValuePair<String, List<StudentObj>> k in 普通科)
            {
                cs[index, 1].PutValue(getDeptCode(k.Key)); //科別代碼
                cs[index, 2].PutValue(k.Key); //科別名稱
                cs[index, 4].PutValue(getClassCount(k.Value)); //班級數
                cs[index, 5].PutValue(getClassCount(k.Value, "1")); //一年級班級數
                cs[index, 6].PutValue(getClassCount(k.Value, "2")); //二年級班級數
                cs[index, 7].PutValue(getClassCount(k.Value, "3")); //三年級班級數
                cs[index, 8].PutValue(getClassCount(k.Value, "4")); //四年級班級數

                cs[index, 9].PutValue(getStudentCount(k.Value)); //總學生數
                cs[index, 10].PutValue(getStudentCount(k.Value, "1")); //總男學生數
                cs[index, 11].PutValue(getStudentCount(k.Value, "0")); //總女學生數
                cs[index, 12].PutValue(getStudentCount(k.Value, "1", "1")); //一年級男學生數
                cs[index, 13].PutValue(getStudentCount(k.Value, "1", "0")); //一年級女學生數
                cs[index, 14].PutValue(getStudentCount(k.Value, "2", "1")); //二年級男學生數
                cs[index, 15].PutValue(getStudentCount(k.Value, "2", "0")); //二年級女學生數
                cs[index, 16].PutValue(getStudentCount(k.Value, "3", "1")); //三年級男學生數
                cs[index, 17].PutValue(getStudentCount(k.Value, "3", "0")); //三年級女學生數
                cs[index, 18].PutValue(getStudentCount(k.Value, "4", "1")); //四年級男學生數
                cs[index, 19].PutValue(getStudentCount(k.Value, "4", "0")); //四年級女學生數
                cs[index, 20].PutValue(getStudentCount(k.Value, "5", "1")); //延修男學生數
                cs[index, 21].PutValue(getStudentCount(k.Value, "5", "0")); //延修女學生數

                cs[index, 22].PutValue(getGraduateStudentCount(k.Key));  //上學年畢業生總數
                cs[index, 23].PutValue(getGraduateStudentCount(k.Key, "1")); //上學年男畢業生總數
                cs[index, 24].PutValue(getGraduateStudentCount(k.Key, "0")); //上學年女畢業生總數

                cs[index, 25].PutValue(getReStudentCount(k.Key));  //當學年重讀生總數
                cs[index, 26].PutValue(getReStudentCount(k.Key, "1"));  //當學年重讀男生總數
                cs[index, 27].PutValue(getReStudentCount(k.Key, "0"));  //當學年重讀女生總數
                cs[index, 28].PutValue(getReStudentCount(k.Key, "1", "1")); //當學年重讀一年級男生總數
                cs[index, 29].PutValue(getReStudentCount(k.Key, "1", "0")); //當學年重讀一年級女生總數
                cs[index, 30].PutValue(getReStudentCount(k.Key, "2", "1")); //當學年重讀二年級男生總數
                cs[index, 31].PutValue(getReStudentCount(k.Key, "2", "0")); //當學年重讀二年級女生總數
                cs[index, 32].PutValue(getReStudentCount(k.Key, "3", "1")); //當學年重讀三年級男生總數
                cs[index, 33].PutValue(getReStudentCount(k.Key, "3", "0")); //當學年重讀三年級女生總數
                index++;
            }

            //綜合高中科
            ws = _Wk.Worksheets[1];
            ws.Name = "綜合高中科";
            cs = ws.Cells;
            index = 9;
            foreach (KeyValuePair<String, List<StudentObj>> k in 綜合高中科)
            {
                cs[index, 1].PutValue(getDeptCode(k.Key)); //科別代碼
                cs[index, 2].PutValue(k.Key); //科別名稱
                cs[index, 4].PutValue(getClassCount(k.Value)); //班級數
                cs[index, 5].PutValue(getClassCount(k.Value, "1")); //一年級班級數
                cs[index, 6].PutValue(getClassCount(k.Value, "2")); //二年級班級數
                cs[index, 7].PutValue(getClassCount(k.Value, "3")); //三年級班級數
                cs[index, 8].PutValue(getClassCount(k.Value, "4")); //四年級班級數

                cs[index, 9].PutValue(getStudentCount(k.Value)); //總學生數
                cs[index, 10].PutValue(getStudentCount(k.Value, "1")); //總男學生數
                cs[index, 11].PutValue(getStudentCount(k.Value, "0")); //總女學生數
                cs[index, 12].PutValue(getStudentCount(k.Value, "1", "1")); //一年級男學生數
                cs[index, 13].PutValue(getStudentCount(k.Value, "1", "0")); //一年級女學生數
                cs[index, 14].PutValue(getStudentCount(k.Value, "2", "1")); //二年級男學生數
                cs[index, 15].PutValue(getStudentCount(k.Value, "2", "0")); //二年級女學生數
                cs[index, 16].PutValue(getStudentCount(k.Value, "3", "1")); //三年級男學生數
                cs[index, 17].PutValue(getStudentCount(k.Value, "3", "0")); //三年級女學生數
                cs[index, 18].PutValue(getStudentCount(k.Value, "4", "1")); //四年級男學生數
                cs[index, 19].PutValue(getStudentCount(k.Value, "4", "0")); //四年級女學生數
                cs[index, 20].PutValue(getStudentCount(k.Value, "5", "1")); //延修男學生數
                cs[index, 21].PutValue(getStudentCount(k.Value, "5", "0")); //延修女學生數

                cs[index, 22].PutValue(getGraduateStudentCount(k.Key));  //上學年畢業生總數
                cs[index, 23].PutValue(getGraduateStudentCount(k.Key, "1")); //上學年男畢業生總數
                cs[index, 24].PutValue(getGraduateStudentCount(k.Key, "0")); //上學年女畢業生總數

                cs[index, 25].PutValue(getReStudentCount(k.Key));  //當學年重讀生總數
                cs[index, 26].PutValue(getReStudentCount(k.Key, "1"));  //當學年重讀男生總數
                cs[index, 27].PutValue(getReStudentCount(k.Key, "0"));  //當學年重讀女生總數
                cs[index, 28].PutValue(getReStudentCount(k.Key, "1", "1")); //當學年重讀一年級男生總數
                cs[index, 29].PutValue(getReStudentCount(k.Key, "1", "0")); //當學年重讀一年級女生總數
                cs[index, 30].PutValue(getReStudentCount(k.Key, "2", "1")); //當學年重讀二年級男生總數
                cs[index, 31].PutValue(getReStudentCount(k.Key, "2", "0")); //當學年重讀二年級女生總數
                cs[index, 32].PutValue(getReStudentCount(k.Key, "3", "1")); //當學年重讀三年級男生總數
                cs[index, 33].PutValue(getReStudentCount(k.Key, "3", "0")); //當學年重讀三年級女生總數
                index++;
            }

            //職業科
            ws = _Wk.Worksheets[2];
            ws.Name = "職業科";
            cs = ws.Cells;
            index = 9;
            foreach (KeyValuePair<String, List<StudentObj>> k in 職業科)
            {
                cs[index, 1].PutValue(getDeptCode(k.Key)); //科別代碼
                cs[index, 2].PutValue(k.Key); //科別名稱
                cs[index, 4].PutValue(getClassCount(k.Value)); //班級數
                cs[index, 5].PutValue(getClassCount(k.Value, "1")); //一年級班級數
                cs[index, 6].PutValue(getClassCount(k.Value, "2")); //二年級班級數
                cs[index, 7].PutValue(getClassCount(k.Value, "3")); //三年級班級數
                cs[index, 8].PutValue(getClassCount(k.Value, "4")); //四年級班級數

                cs[index, 9].PutValue(getStudentCount(k.Value)); //總學生數
                cs[index, 10].PutValue(getStudentCount(k.Value, "1")); //總男學生數
                cs[index, 11].PutValue(getStudentCount(k.Value, "0")); //總女學生數
                cs[index, 12].PutValue(getStudentCount(k.Value, "1", "1")); //一年級男學生數
                cs[index, 13].PutValue(getStudentCount(k.Value, "1", "0")); //一年級女學生數
                cs[index, 14].PutValue(getStudentCount(k.Value, "2", "1")); //二年級男學生數
                cs[index, 15].PutValue(getStudentCount(k.Value, "2", "0")); //二年級女學生數
                cs[index, 16].PutValue(getStudentCount(k.Value, "3", "1")); //三年級男學生數
                cs[index, 17].PutValue(getStudentCount(k.Value, "3", "0")); //三年級女學生數
                cs[index, 18].PutValue(getStudentCount(k.Value, "4", "1")); //四年級男學生數
                cs[index, 19].PutValue(getStudentCount(k.Value, "4", "0")); //四年級女學生數
                cs[index, 20].PutValue(getStudentCount(k.Value, "5", "1")); //延修男學生數
                cs[index, 21].PutValue(getStudentCount(k.Value, "5", "0")); //延修女學生數

                cs[index, 22].PutValue(getGraduateStudentCount(k.Key));  //上學年畢業生總數
                cs[index, 23].PutValue(getGraduateStudentCount(k.Key, "1")); //上學年男畢業生總數
                cs[index, 24].PutValue(getGraduateStudentCount(k.Key, "0")); //上學年女畢業生總數

                cs[index, 25].PutValue(getReStudentCount(k.Key));  //當學年重讀生總數
                cs[index, 26].PutValue(getReStudentCount(k.Key, "1"));  //當學年重讀男生總數
                cs[index, 27].PutValue(getReStudentCount(k.Key, "0"));  //當學年重讀女生總數
                cs[index, 28].PutValue(getReStudentCount(k.Key, "1", "1")); //當學年重讀一年級男生總數
                cs[index, 29].PutValue(getReStudentCount(k.Key, "1", "0")); //當學年重讀一年級女生總數
                cs[index, 30].PutValue(getReStudentCount(k.Key, "2", "1")); //當學年重讀二年級男生總數
                cs[index, 31].PutValue(getReStudentCount(k.Key, "2", "0")); //當學年重讀二年級女生總數
                cs[index, 32].PutValue(getReStudentCount(k.Key, "3", "1")); //當學年重讀三年級男生總數
                cs[index, 33].PutValue(getReStudentCount(k.Key, "3", "0")); //當學年重讀三年級女生總數
                index++;
            }
        }

        //複製範本
        private Worksheet Template()
        {
            Workbook wk2 = new Workbook();
            wk2.Open(new MemoryStream(Properties.Resources.班級學生概況1));
            return wk2.Worksheets[0];
        }

        //確認資料性別是否有異常
        private void CleanList(List<StudentObj> studentlist)
        {
            _ErrorList = new List<StudentObj>();
            _CorrectList = new List<StudentObj>();
            foreach (StudentObj obj in studentlist)
            {
                if ((obj.性別 != "0" && obj.性別 != "1") || (obj.年級 == ""))
                {
                    _ErrorList.Add(obj);
                }
                else
                {
                    _CorrectList.Add(obj);
                }
            }
        }

        //按照科別取得字典,key=科別名稱,value=學生物件清單
        private Dictionary<String, List<StudentObj>> getDicByDept(Dictionary<String, List<StudentObj>> deptDic, String deptName)
        {
            Dictionary<String, List<StudentObj>> dic = new Dictionary<string, List<StudentObj>>();

            if (deptName == "普通科" || deptName == "綜合高中科")
            {
                foreach (KeyValuePair<String, List<StudentObj>> k in deptDic)
                {
                    if (k.Key.Contains(deptName))
                    {
                        foreach (StudentObj student in k.Value)
                        {
                            if (!dic.ContainsKey(deptName))
                            {
                                dic.Add(deptName, new List<StudentObj>());
                            }
                            dic[deptName].Add(student);
                        }
                    }
                }
            }
            else
            {
                foreach (KeyValuePair<String, List<StudentObj>> k in deptDic)
                {
                    if (!k.Key.Contains("普通科") && !k.Key.Contains("綜合高中科"))
                    {
                        foreach (StudentObj student in k.Value)
                        {
                            if (!dic.ContainsKey(student.科別))
                            {
                                dic.Add(student.科別, new List<StudentObj>());
                            }
                            dic[student.科別].Add(student);
                        }
                    }
                }
            }
            return dic;
        }

        //建立科別代號表
        public void QueryDeptCode()
        {
            Dept_ref = new Dictionary<string, string>();
            QueryHelper _Q = new QueryHelper();
            DataTable dt = _Q.Select("select id,code,name from dept");
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String code = row["code"].ToString();
                if (code == "") code = "NoCode";
                String name = row["name"].ToString();
                Dept_ref.Add(id + "_" + code, name);
            }
        }

        //查詢科別代碼
        public String getDeptCode(String name)
        {
            String code = "";
            foreach (KeyValuePair<String, String> dept_ref in Dept_ref)
            {
                if (name == dept_ref.Value)
                {
                    code = dept_ref.Key.Split('_')[1];
                }
            }
            return code;
        }

        ////字典內容轉換成學生物件清單
        //private List<StudentObj> toAllList(Dictionary<String, List<StudentObj>> dic)
        //{
        //    List<StudentObj> list = new List<StudentObj>();
        //    foreach (KeyValuePair<String, List<StudentObj>> k in dic)
        //    {
        //        foreach (StudentObj student in k.Value)
        //        {
        //            list.Add(student);
        //        }
        //    }
        //    return list;
        //}

        //傳入學生清單取得班級數量
        private int getClassCount(List<StudentObj> list)
        {
            Dictionary<string, List<StudentObj>> dic = new Dictionary<string, List<StudentObj>>();
            foreach (StudentObj student in list)
            {
                if (!dic.ContainsKey(student.班級ID))
                {
                    dic.Add(student.班級ID, new List<StudentObj>());
                }
                dic[student.班級ID].Add(student);
            }
            return dic.Count;
        }

        //傳入學生清單取得指定年級的班級數量
        private int getClassCount(List<StudentObj> list, String grade)
        {
            Dictionary<string, List<StudentObj>> dic = new Dictionary<string, List<StudentObj>>();
            foreach (StudentObj student in list)
            {
                if (student.年級 == grade)
                {
                    if (!dic.ContainsKey(student.班級ID))
                    {
                        dic.Add(student.班級ID, new List<StudentObj>());
                    }
                    dic[student.班級ID].Add(student);
                }
            }
            return dic.Count;
        }

        //傳入學生清單取得學生數量
        private int getStudentCount(List<StudentObj> list)
        {
            int count = 0;
            foreach (StudentObj student in list)
            {
                if (student.狀態 == "1" || student.狀態 == "2")
                {
                    count++;
                }
            }
            return count;
        }

        //傳入學生清單取得指定性別的學生數量
        private int getStudentCount(List<StudentObj> list, String gender)
        {
            int count = 0;
            foreach (StudentObj student in list)
            {
                if (student.性別 == gender && (student.狀態 == "1" || student.狀態 == "2"))
                {
                    count++;
                }
            }
            return count;
        }

        //傳入學生清單取得指定年級,性別的學生數量,grade=5代表查詢延修生
        private int getStudentCount(List<StudentObj> list, String grade, String gender)
        {
            int count = 0;
            switch (grade)
            {
                case "5":
                    String status = "2";
                    foreach (StudentObj student in list)
                    {
                        if (student.狀態 == status && student.性別 == gender)
                        {
                            count++;
                        }
                    }
                    break;

                default:
                    foreach (StudentObj student in list)
                    {
                        if (student.年級 == grade && student.性別 == gender && student.狀態 == "1")
                        {
                            count++;
                        }
                    }
                    break;
            }
            return count;
        }

        //取得指定科別的畢業生總數
        private int getGraduateStudentCount(String deptName)
        {
            int count = 0;
            foreach (GraduateStudentObj student in _GraduateStudentList)
            {
                if (student.科別.Contains(deptName))
                {
                    count++;
                }
            }
            return count;
        }

        //取得指定科別及性別的畢業生總數
        private int getGraduateStudentCount(String deptName, String gender)
        {
            int count = 0;
            foreach (GraduateStudentObj student in _GraduateStudentList)
            {
                if (student.科別.Contains(deptName) && student.性別 == gender)
                {
                    count++;
                }
            }
            return count;
        }

        //取得指定科別的重讀生數量
        private int getReStudentCount(String deptName)
        {
            int count = 0;
            foreach (ReStudentObj student in _ReStudentList)
            {
                if (student.科別.Contains(deptName))
                {
                    count++;
                }
            }
            return count;
        }

        //取得指定科別及性別的重讀生數量
        private int getReStudentCount(String deptName, String gender)
        {
            int count = 0;
            foreach (ReStudentObj student in _ReStudentList)
            {
                if (student.科別.Contains(deptName) && student.性別 == gender)
                {
                    count++;
                }
            }
            return count;
        }

        //取得指定科別年級及性別的重讀生總數
        private int getReStudentCount(String deptName, String grade, String gender)
        {
            int count = 0;
            foreach (ReStudentObj student in _ReStudentList)
            {
                if (student.科別.Contains(deptName) && student.年級 == grade && student.性別 == gender)
                {
                    count++;
                }
            }
            return count;
        }

        //取得上學年畢業生物件清單
        private List<GraduateStudentObj> getGraduateStudent()
        {
            int year = Convert.ToInt32(_SchoolYear)-1; //當前系統學年度-1
            List<GraduateStudentObj> list = new List<GraduateStudentObj>();
            FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();
            DataTable dt = _Q.Select("select ref_student_id,ss_name,ss_gender,ss_dept from update_record where update_code='501' and school_year='" + year + "'");

            foreach (DataRow row in dt.Rows)
            {
                list.Add(new GraduateStudentObj(row));
            }
            return list;
        }

        //取得重讀學生物件清單
        private List<ReStudentObj> getReStudent()
        {
            List<ReStudentObj> list = new List<ReStudentObj>();
            FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();
            DataTable dt = _Q.Select("select ref_student_id,ss_name,ss_gender,ss_grade_year,ss_dept from update_record where school_year='" + _SchoolYear + "' and update_desc like '%重讀%'");

            foreach (DataRow row in dt.Rows)
            {
                list.Add(new ReStudentObj(row));
            }
            return list;
        }
    }
}
