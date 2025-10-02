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

        //        private List<CompletionStudentObj> _CompletionStudentList;
        //private List<ReStudentObj> _ReStudentList;

        // 一般生、延修生:取得學生來源：學生狀態2+異動代碼:235+異動年級：延修生。
        private Dictionary<String, List<StudentObj>> DeptDic;


        private Dictionary<String, List<StudentObj>> DeptDic2;

        Dictionary<String, String> Dept_ref; //科別代碼對照,key=code,value=name;

        private BackgroundWorker _BGWClassStudentAbsenceDetail; //背景模式

        Workbook _Wk;
        List<string> Public_BranchIDs;

        string Public_BranchName;
        Boolean HasData;
        public void Start(List<string> BranchIDs, string BranchName)
        {
            Public_BranchIDs = BranchIDs;
            Public_BranchName = BranchName;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在產生班級及學生概況統計表...");
            _BGWClassStudentAbsenceDetail = new BackgroundWorker();
            _BGWClassStudentAbsenceDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentAbsenceDetail_DoWork);
            _BGWClassStudentAbsenceDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWClassStudentAbsenceDetail_Completed);
            _BGWClassStudentAbsenceDetail.RunWorkerAsync();
        }

        private void _BGWClassStudentAbsenceDetail_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            if (HasData)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 班級及學生概況統計表 已完成");
                SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = "班級及學生概況統計表.xlsx";
                sd.Filter = "Excel檔案 (*.xlsx)|*.xls|所有檔案 (*.*)|*.*";
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
            else
                FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 班級及學生概況統計表 已完成");

        }


        private void _BGWClassStudentAbsenceDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            _SchoolYear = K12.Data.School.DefaultSchoolYear;
            _GraduateStudentList = getGraduateStudent();

            QueryDeptCode(); //建立科別代號表

            FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();
            // 一般生、延修生:取得學生來源：學生狀態2+異動代碼:235+異動年級：延修生。
            string sql = string.Format(@"
            WITH student_data AS (
               SELECT
                student.id,
                student.name,
                student.ref_class_id,
                gender,
                class.grade_year,
                student.status,
                COALESCE(student.ref_dept_id, class.ref_dept_id) AS dept_id
            FROM
                student
                LEFT JOIN class ON student.ref_class_id = class.id
            WHERE
                student.status = 1
                AND class.grade_year IN(1, 2, 3)
            UNION
            ALL
            SELECT
                student.id,
                student.name,
                student.ref_class_id,
                gender,
                class.grade_year,
                student.status,
                COALESCE(student.ref_dept_id, class.ref_dept_id) AS dept_id
            FROM
                student
                LEFT JOIN class ON student.ref_class_id = class.id
            WHERE
                student.id IN(
                    SELECT
                        student.id AS student_id
                    FROM
                        student
                        INNER JOIN update_record ON student.id = update_record.ref_student_id
                    WHERE
                        student.status = 2
                        AND update_code = '235'
                        AND ss_grade_year = -1
                )
            )
            SELECT
                student_data.*,
                dept.code AS dept_code,
                dept.name AS dept_name,
                dept.ref_dept_group_id
            FROM
                student_data
                INNER JOIN dept ON student_data.dept_id = dept.id
                AND dept.ref_dept_group_id IN({0});
            ", string.Join(",", Public_BranchIDs.ToArray()));

            DataTable dt = _Q.Select(sql);

            List<StudentObj> StudentList = new List<StudentObj>();
            if (dt.Rows.Count > 0 || _GraduateStudentList.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    StudentObj obj = new StudentObj(row);
                    StudentList.Add(obj);
                }

                //資料整理
                CleanList(StudentList);
                DeptDic = SortToDept(_CorrectList); //科別分類
                                                    //普通科 = getDicByDept(DeptDic, "普通科");
                                                    //綜合高中科 = getDicByDept(DeptDic, "綜合高中科");
                                                    //職業科 = getDicByDept(DeptDic, "職業科");

                //資料列印           
                Export();
                HasData = true;
            }
            else
            {
                MessageBox.Show("該部門沒有科別有資料");
                HasData = false;

            }
        }

        //科別分類,key=科別名稱,value=學生物件清單
        private Dictionary<String, List<StudentObj>> SortToDept(List<StudentObj> StudentList)
        {
            Dictionary<String, List<StudentObj>> Dic = new Dictionary<string, List<StudentObj>>();
            foreach (StudentObj obj in StudentList)
            {
                if (!Dic.ContainsKey(obj.Department)) //建立科別
                {
                    Dic.Add(obj.Department, new List<StudentObj>());
                }
                Dic[obj.Department].Add(obj);
            }
            return Dic;
        }



        private void Export()
        {
            _Wk = new Workbook();
            _Wk.Worksheets.Add();
            _Wk.Worksheets[0].Copy(Template()); //複製範本

            Worksheet ws = _Wk.Worksheets[1];
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
                cs[index, 0].PutValue(obj.StudentID);
                cs[index, 1].PutValue(obj.StudentName);
                cs[index, 2].PutValue(obj.ClassID);
                cs[index, 3].PutValue(obj.Gender);
                cs[index, 4].PutValue(obj.GradeYear);
                cs[index, 5].PutValue(obj.Department);
                index++;
            }

            //指定部別
            ws = _Wk.Worksheets[0];
            ws.Name = Public_BranchName;
            cs = ws.Cells;
            index = 9;
            cs[2, 0].PutValue("表-7 高中職學校班級及學生概況（一）－" + Public_BranchName);
            cs[3, 12].PutValue(_SchoolYear);
            cs[4, 0].PutValue(K12.Data.School.Code);
            cs[4, 2].PutValue(Public_BranchName);
            cs[0, 29].PutValue(K12.Data.School.ChineseName + "(教務處)");
            string DeptName = "";
            foreach (string deptName in DeptDic.Keys)
            {
                List<StudentObj> studentList = DeptDic[deptName];
                //cs[index, 1].PutValue(getDeptCode(k.Key)); //科別代碼
                //cs[index, 2].PutValue(k.Key); //科別名稱
                cs[index, 1].PutValue(getDeptCode(deptName)); //科別代碼
                cs[index, 2].PutValue(deptName); //科別名稱
                cs[index, 4].PutValue(getClassCount(studentList)); //班級數
                cs[index, 5].PutValue(getClassCount(studentList, "1")); //一年級班級數
                cs[index, 6].PutValue(getClassCount(studentList, "2")); //二年級班級數
                cs[index, 7].PutValue(getClassCount(studentList, "3")); //三年級班級數 
                cs[index, 8].PutValue(getStudentCount(studentList)); //總學生數
                cs[index, 9].PutValue(getStudentCount(studentList, "1")); //總男學生數
                cs[index, 10].PutValue(getStudentCount(studentList, "0")); //總女學生數
                cs[index, 11].PutValue(getStudentCount(studentList, "1", "1")); //一年級男學生數
                cs[index, 12].PutValue(getStudentCount(studentList, "1", "0")); //一年級女學生數
                cs[index, 13].PutValue(getStudentCount(studentList, "2", "1")); //二年級男學生數
                cs[index, 14].PutValue(getStudentCount(studentList, "2", "0")); //二年級女學生數
                cs[index, 15].PutValue(getStudentCount(studentList, "3", "1")); //三年級男學生數
                cs[index, 16].PutValue(getStudentCount(studentList, "3", "0")); //三年級女學生數 
                cs[index, 17].PutValue(getStudentCount(studentList, "4", "1")); //延修男學生數
                cs[index, 18].PutValue(getStudentCount(studentList, "4", "0")); //延修女學生數

                cs[index, 19].PutValue(getGraduateStudentCount(deptName));  //上學年畢業生總數
                cs[index, 20].PutValue(getGraduateStudentCount(deptName, "1")); //上學年男畢業生總數
                cs[index, 21].PutValue(getGraduateStudentCount(deptName, "0")); //上學年女畢業生總數

                // 這些新版2025不需要呈報先註解
                //cs[index, 22].PutValue(getCompletionStudentCount(k.Key));  //上學年修業生總數                
                //cs[index, 23].PutValue(getCompletionStudentCount(k.Key, "1")); //上學年男修業生總數
                //cs[index, 24].PutValue(getCompletionStudentCount(k.Key, "0")); //上學年女修業生總數

                //cs[index, 25].PutValue(getReStudentCount(k.Key));  //當學年重讀生總數
                //cs[index, 26].PutValue(getReStudentCount(k.Key, "1"));  //當學年重讀男生總數
                //cs[index, 27].PutValue(getReStudentCount(k.Key, "0"));  //當學年重讀女生總數
                //cs[index, 28].PutValue(getReStudentCount(k.Key, "1", "1")); //當學年重讀一年級男生總數
                //cs[index, 29].PutValue(getReStudentCount(k.Key, "1", "0")); //當學年重讀一年級女生總數
                //cs[index, 30].PutValue(getReStudentCount(k.Key, "2", "1")); //當學年重讀二年級男生總數
                //cs[index, 31].PutValue(getReStudentCount(k.Key, "2", "0")); //當學年重讀二年級女生總數
                //cs[index, 32].PutValue(getReStudentCount(k.Key, "3", "1")); //當學年重讀三年級男生總數
                //cs[index, 33].PutValue(getReStudentCount(k.Key, "3", "0")); //當學年重讀三年級女生總數
                index++;
            }

        }

        //複製範本
        private Worksheet Template()
        {
            Workbook wk2 = new Workbook();
            wk2.Open(new MemoryStream(Properties.Resources.班級學生概況2025));
            return wk2.Worksheets[0];
        }

        //確認資料性別是否有異常
        private void CleanList(List<StudentObj> studentlist)
        {
            _ErrorList = new List<StudentObj>();
            _CorrectList = new List<StudentObj>();
            foreach (StudentObj obj in studentlist)
            {
                if ((obj.Gender != "0" && obj.Gender != "1") || (obj.GradeYear == ""))
                {
                    _ErrorList.Add(obj);
                }
                else
                {
                    _CorrectList.Add(obj);
                }
            }
        }


        //建立科別代號表
        public void QueryDeptCode()
        {
            try
            {
                Dept_ref = new Dictionary<string, string>();
                QueryHelper _Q = new QueryHelper();
                string sql = string.Format(@"
            SELECT
                code,
                name
            FROM
                dept;
            ");
                DataTable dt = _Q.Select(sql);
                foreach (DataRow row in dt.Rows)
                {
                    // 科別代碼
                    string code = row["code"].ToString();
                    // 科別名稱，是唯一值，介面有檔。
                    string name = row["name"].ToString();
                    if (!Dept_ref.ContainsKey(name))
                    {
                        Dept_ref.Add(name, code);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("QueryDeptCode," + ex.Message);
            }

        }

        //查詢科別代碼
        public string getDeptCode(string name)
        {
            string code = "";
            if (Dept_ref.ContainsKey(name))
            {
                code = Dept_ref[name].ToString();
            }
            return code;
        }


        //傳入學生清單取得班級數量
        private int getClassCount(List<StudentObj> list)
        {
            if (list == null || list.Count == 0)
                return 0;

            // 使用 HashSet 來儲存唯一的 ClassID，效能更好
            HashSet<string> uniqueClassIDs = new HashSet<string>();
            foreach (StudentObj student in list)
            {
                // 只計算一般狀態
                if (student.Status == "1")
                    if (!string.IsNullOrEmpty(student.ClassID))
                    {
                        uniqueClassIDs.Add(student.ClassID);
                    }
            }
            return uniqueClassIDs.Count;
        }

        //傳入學生清單取得指定年級的班級數量
        private int getClassCount(List<StudentObj> list, String grade)
        {
            if (list == null || list.Count == 0 || string.IsNullOrEmpty(grade))
                return 0;

            // 使用 HashSet 來儲存唯一的 ClassID，效能更好
            HashSet<string> uniqueClassIDs = new HashSet<string>();
            foreach (StudentObj student in list)
            {
                if (student.Status == "1")
                    if (!string.IsNullOrEmpty(student.ClassID) && student.GradeYear == grade)
                    {
                        uniqueClassIDs.Add(student.ClassID);
                    }
            }
            return uniqueClassIDs.Count;
        }

        //傳入學生清單取得學生數量
        private int getStudentCount(List<StudentObj> list)
        {
            return list.Count;
        }

        //傳入學生清單取得指定性別的學生數量
        private int getStudentCount(List<StudentObj> list, String gender)
        {
            // 使用 LINQ 和 HashSet 優化
            var validGrades = new HashSet<string> { "1", "2", "3" };
            return list.Count(student => student.Gender == gender && validGrades.Contains(student.GradeYear));
        }

        //傳入學生清單取得指定年級,性別的學生數量,grade=5代表查詢延修生
        private int getStudentCount(List<StudentObj> list, String grade, String gender)
        {
            if (grade == "4")
                return list.Count(student => student.Gender == gender && student.Status == "2");
            else
                return list.Count(student => student.GradeYear == grade && student.Gender == gender && student.Status == "1");
        }

        //取得指定科別的畢業生總數
        private int getGraduateStudentCount(string deptName)
        {
            // 使用 LINQ 優化，效能更好且程式碼更簡潔
            return _GraduateStudentList.Count(student => student.Department == deptName);
        }

        //取得指定科別及性別的畢業生總數
        private int getGraduateStudentCount(string deptName, string gender)
        {
            // 使用 LINQ 優化，效能更好且程式碼更簡潔
            return _GraduateStudentList.Count(student =>
                student.Department == deptName && student.Gender == gender);
        }

        ////取得指定科別的修業生總數
        //private int getCompletionStudentCount(String deptName)
        //{
        //    int count = 0;
        //    foreach (CompletionStudentObj student in _CompletionStudentList)
        //    {
        //        if (student.DepartmentID + "_" + student.Department == deptName)
        //        {
        //            count++;
        //        }
        //    }
        //    return count;
        //}

        //取得指定科別及性別的修業生總數
        //private int getCompletionStudentCount(String deptName, String gender)
        //{
        //    int count = 0;
        //    foreach (CompletionStudentObj student in _CompletionStudentList)
        //    {
        //        if (student.DepartmentID + "_" + student.Department == deptName && student.Gender == gender)
        //        {
        //            count++;
        //        }
        //    }
        //    return count;
        //}
        //取得指定科別的重讀生數量
        //private int getReStudentCount(String deptName)
        //{
        //    int count = 0;
        //    foreach (ReStudentObj student in _ReStudentList)
        //    {
        //        if (student.DepartmentID + "_" + student.Department == deptName)
        //        {
        //            count++;
        //        }
        //    }
        //    return count;
        //}

        //取得指定科別及性別的重讀生數量
        //private int getReStudentCount(String deptName, String gender)
        //{
        //    int count = 0;
        //    foreach (ReStudentObj student in _ReStudentList)
        //    {
        //        if (student.DepartmentID + "_" + student.Department == deptName && student.Gender == gender)
        //        {
        //            count++;
        //        }
        //    }
        //    return count;
        //}

        //取得指定科別年級及性別的重讀生總數
        //private int getReStudentCount(String deptName, String grade, String gender)
        //{
        //    int count = 0;
        //    foreach (ReStudentObj student in _ReStudentList)
        //    {
        //        if (student.DepartmentID + "_" + student.Department == deptName && student.GradeYear == grade && student.Gender == gender)
        //        {
        //            count++;
        //        }
        //    }
        //    return count;
        //}

        //取得上學年畢業生物件清單
        private List<GraduateStudentObj> getGraduateStudent()
        {
            List<GraduateStudentObj> list = new List<GraduateStudentObj>();

            try
            {
                int year = Convert.ToInt32(_SchoolYear) - 1; //當前系統學年度-1            
                FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();

                string sql = string.Format(@"
                WITH student_data AS (
                    SELECT
                        update_record.ref_student_id,
                        update_record.ss_name,
                        update_record.ss_gender,
                        update_record.ss_dept,
                        COALESCE(student.ref_dept_id, class.ref_dept_id) AS dept_id,
                        school_year
                    FROM
                        student
                        LEFT JOIN class ON student.ref_class_id = class.id
                        INNER JOIN update_record ON update_record.ref_student_id = student.id
                    WHERE
                        student.status = 16
                        AND update_record.update_code = '501'
                        AND update_record.school_year = {0}
                )
                SELECT
                    student_data.*,
                    dept.code AS dept_code,
                    dept.name AS dept_name,
                    dept.ref_dept_group_id
                FROM
                    student_data
                    JOIN dept ON student_data.dept_id = dept.id
                WHERE
                    dept.ref_dept_group_id IN({1});
                ", year, string.Join(",", Public_BranchIDs.ToArray()));

                DataTable dt = _Q.Select(sql);

                foreach (DataRow row in dt.Rows)
                {
                    list.Add(new GraduateStudentObj(row));
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("getGraduateStudent()," + e.Message);
            }

            return list;
        }
        //取得上學年修業生物件清單
        //private List<CompletionStudentObj> getCompletionStudent()
        //{
        //    int year = Convert.ToInt32(_SchoolYear) - 1; //當前系統學年度-1
        //    List<CompletionStudentObj> list = new List<CompletionStudentObj>();
        //    FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();
        //    string sql = @"WITH student_data AS (
        //                 SELECT update_record.ref_student_id, update_record.ss_name, update_record.ss_gender
        //                 , update_record.ss_dept, COALESCE(student.ref_dept_id,class.ref_dept_id ) AS _dept
        //                 , update_code, school_year
        //                 FROM student LEFT JOIN class ON student.ref_class_id=class.id
        //                      LEFT JOIN update_record ON  update_record.ref_student_id=student.id)
        //            SELECT
        //                    student_data.*, dept.name AS dept_name,dept.ref_dept_group_id
        //                FROM
        //                    student_data  JOIN  dept ON student_data._dept= dept.id ";
        //    if (chkUnGraduate)
        //        sql = sql + " where update_code in ('366','364')";
        //    else
        //        sql = sql + " where update_code in ('366')";
        //    sql = sql + " and school_year = " + year;
        //    sql = sql + " AND dept.ref_dept_group_id in (" + Public_BranchID.Substring(0, Public_BranchID.Length - 1) + ")";
        //    DataTable dt = _Q.Select(sql);

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        list.Add(new CompletionStudentObj(row));
        //    }
        //    return list;
        //}
        //取得重讀學生物件清單
        //        private List<ReStudentObj> getReStudent()
        //        {
        //            List<ReStudentObj> list = new List<ReStudentObj>();
        //            FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();
        //            DateTime dt_C = DateTime.Parse((Convert.ToInt32(_SchoolYear) + 1911).ToString() + "/10/01");
        //            string sql = @"WITH student_data AS (
        //                         SELECT update_record.ref_student_id, update_record.ss_name, update_record.ss_gender
        //                         , update_record.ss_dept, 
        //             update_record.update_date,update_record.ss_grade_year,
        //COALESCE(student.ref_dept_id,class.ref_dept_id ) AS _dept
        //                         , update_code, school_year
        //                         FROM student LEFT JOIN class ON student.ref_class_id=class.id
        //                              LEFT JOIN update_record ON  update_record.ref_student_id=student.id)
        //                    SELECT
        //                            student_data.*, dept.name AS dept_name,dept.ref_dept_group_id
        //                        FROM
        //                            student_data  JOIN  dept ON student_data._dept= dept.id 
        //                 where school_year = " + _SchoolYear + " and update_code in ('231','232','233','234','237', '238','239','240','241','242') " + " and update_date< '" + dt_C.ToShortDateString() + "'";
        //            sql = sql + " AND  dept.ref_dept_group_id in (" + Public_BranchID.Substring(0, Public_BranchID.Length - 1) + ")";
        //            DataTable dt = _Q.Select(sql);

        //            foreach (DataRow row in dt.Rows)
        //            {
        //                list.Add(new ReStudentObj(row));
        //            }
        //            return list;
        //        }
    }
}
