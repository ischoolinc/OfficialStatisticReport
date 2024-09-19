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
using ArrangeClass.DAO;

namespace ArrangeClass
{
    class Printer
    {
        //本案為2017/12/5 穎驊因應 教育部國民及學前教育署 的公文而作，
        //其 暨大的計畫 "全國高級中等學校學生基本資料庫 新增了 "編班名冊"報表 上傳，
        // 每年11/1 、 4/1 需上傳，詳細規格內容 可見 Resources 資源檔

        private List<string> _ErrorList, _CorrectList;

        //取得班級名稱對照設定檔
        Dictionary<string, string> _ClassNameMappingDict = QueryTransfer.GetConfigure();

        //有錯誤不完整資料的學生
        List<SHSchool.Data.SHStudentRecord> error_StudentList = new List<SHSchool.Data.SHStudentRecord>();

        FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "資料統計"];

        Dictionary<String, String> Dept_ref = new Dictionary<string, string>(); //科別代碼對照,key=name,value=code;
        private BackgroundWorker _BGWClassStudentAbsenceDetail; //背景模式
        Workbook _Wk;

        public void Start()
        {
            //由於本功能報表需要一段時間產生，先關閉按鈕功能，怕使用者連續點取兩下            
            item1["報表"]["編班名冊"].Enable = false;

            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在產生編班名冊...", 0);
            _BGWClassStudentAbsenceDetail = new BackgroundWorker();
            _BGWClassStudentAbsenceDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentAbsenceDetail_DoWork);
            _BGWClassStudentAbsenceDetail.WorkerReportsProgress = true;

            _BGWClassStudentAbsenceDetail.ProgressChanged += delegate (object vsender, ProgressChangedEventArgs ve)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在產生編班名冊...", ve.ProgressPercentage);
            };

            _BGWClassStudentAbsenceDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWClassStudentAbsenceDetail_Completed);
            _BGWClassStudentAbsenceDetail.RunWorkerAsync();
        }

        private void _BGWClassStudentAbsenceDetail_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            //重啟按鈕功能
            item1["報表"]["編班名冊"].Enable = true;


            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 編班名冊 已完成");
            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "編班名冊.xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    //_Wk.Save(sd.FileName);
                    _Wk.Save(sd.FileName, SaveFormat.Excel97To2003);  // 2021-03-16 要求上傳檔案為xls

                    if (error_StudentList.Count > 0)
                    {
                        MessageBox.Show("發現" + error_StudentList.Count + "筆異常資料未列入統計\r\n詳細資料請確認報表中的[錯誤報告]");
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
            //建立科別代碼對照
            QueryDeptCode();

            _Wk = new Workbook(new MemoryStream(Properties.Resources.編班名冊欄位範例20171114_樣板_));

            //全校所有的學生
            List<SHSchool.Data.SHStudentRecord> all_StudentList = SHSchool.Data.SHStudent.SelectAll();
            //整理後的學生
            List<SHSchool.Data.SHStudentRecord> target_StudentList = new List<SHSchool.Data.SHStudentRecord>();

            //學生資料錯誤的原因
            Dictionary<string, string> errorReasonDict = new Dictionary<string, string>();

            // 濾出 狀態 為 一般、休學、延修 的學生
            foreach (SHSchool.Data.SHStudentRecord sr in all_StudentList)
            {
                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.一般 || sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.休學 || sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.延修)
                {
                    target_StudentList.Add(sr);

                }
            }

            _BGWClassStudentAbsenceDetail.ReportProgress(30);

            Worksheet ws = _Wk.Worksheets[0];

            Worksheet ws_errorReport = _Wk.Worksheets[1];

            Cells cs = ws.Cells;

            Cells cs_errorReport = ws_errorReport.Cells;

            int row_counter = 1;

            string schoolCode = K12.Data.School.Code;

            string schoolYear = K12.Data.School.DefaultSchoolYear;

            string schoolSemester = K12.Data.School.DefaultSemester;

            // 本編班名冊別  固定為c
            string bookType = "c";
            //身分別
            string identityType = "";
            //年級
            string grade = "";
            //班級名稱
            string className = "";
            //座號
            string seatNo = "";

            //班別
            //※高中、高職：
            //日間部填1、夜間部填2、實用技能學程填3、建教班填4、重點產業班 / 台德菁英班 / 雙軌旗艦訓練計畫專班填7、建教僑生專班填8
            //※進修部(學校)：
            //核定班填01、員工進修班填04、重點產業班填05、產業人力套案專班填06
            string classCode = "";

            //科別學程代碼
            string departmentCode = "";

            //上傳類別
            string updateType = "";

            //錯誤資料不齊全原因
            string errorReason = "";

            //取得所有目標學生的異動資料
            List<SHUpdateRecordRecord> updateList = SHSchool.Data.SHUpdateRecord.SelectByStudents(target_StudentList);

            Dictionary<string, List<SHUpdateRecordRecord>> updateDict = new Dictionary<string, List<SHUpdateRecordRecord>>();

            _BGWClassStudentAbsenceDetail.ReportProgress(40);

            List<string> StudentIDs = new List<string>();

            // 整理學生全部異動資料，以利後續對出學生班別
            foreach (SHUpdateRecordRecord srr in updateList)
            {
                StudentIDs.Add(srr.StudentID);

                if (!updateDict.ContainsKey(srr.StudentID))
                {
                    updateDict.Add(srr.StudentID, new List<SHUpdateRecordRecord>());

                    updateDict[srr.StudentID].Add(srr);
                }
                else
                {
                    updateDict[srr.StudentID].Add(srr);
                }
            }

            // 取得學生學期對照班級學生
            Dictionary<string, List<SemsHistoryInfo>> StudSemsHisDict = QueryTransfer.GetSemsHistoryInfoByIDs(StudentIDs);


            // 上傳類別的對應 Dictionary ，使用班別來對照
            Dictionary<string, string> updateTypeDict = new Dictionary<string, string>();

            // 科別名稱與代碼對照
            Dictionary<string, string> deptNameCodeDict = QueryTransfer.GetDeptNameCodeDict();

            //日間部
            updateTypeDict.Add("1", "a");
            //夜間部
            updateTypeDict.Add("2", "b");
            //實用技能學程(日間)
            updateTypeDict.Add("3", "c");
            //重點產業班
            updateTypeDict.Add("7", "e");

            //目前2017/12/7， 以下 類別 無法單純從班別來對照出來，需要使用者手動補齊，目前程式碼僅能就大部分的情況將就
            //實用技能學程(日間)請填c、實用技能學程(夜間)請填d、重點產業班請填e、台德菁英班/雙軌旗艦訓練計畫專班請填f、雙語部填h

            //進修部 代碼都是 g          
            updateTypeDict.Add("01", "g");
            updateTypeDict.Add("04", "g");
            updateTypeDict.Add("05", "g");
            updateTypeDict.Add("06", "g");

            // 取得 教務作業>批次作業/檢視>異動作業>核班人數維護 資料內容，加入使用者設定
            Dictionary<string, string> ClasssTypeU = QueryTransfer.GetClassTyepUDict(K12.Data.School.DefaultSchoolYear);
            foreach (string key in ClasssTypeU.Keys)
            {
                if (!updateTypeDict.ContainsKey(key))
                    updateTypeDict.Add(key, ClasssTypeU[key]);
            }

            _BGWClassStudentAbsenceDetail.ReportProgress(50);

            #region 報表填值
            //填寫資料
            foreach (SHSchool.Data.SHStudentRecord sr in target_StudentList)
            {
                identityType = "";
                grade = "";
                className = "";
                seatNo = "";
                departmentCode = "";
                updateType = "";
                errorReason = "";

                // 最後一筆異動
                SHUpdateRecordRecord lastUpdateRec = null;

                // 班別取得規則異動最後一筆有核准日期或臨編日期
                if (updateDict.ContainsKey(sr.ID))
                {
                    // 有核准日期或臨編日期放入
                    foreach (SHUpdateRecordRecord srr in updateDict[sr.ID])
                    {
                        // 處理核准日期
                        if (srr.ADDate != "")
                        {
                            DateTime d1;
                            if (DateTime.TryParse(srr.ADDate, out d1))
                            {
                                if (lastUpdateRec == null)
                                {
                                    lastUpdateRec = srr;
                                }
                                else
                                {
                                    DateTime d2;
                                    if (DateTime.TryParse(lastUpdateRec.ADDate, out d2))
                                    {
                                        if (d1 > d2)
                                        {
                                            lastUpdateRec = srr;
                                        }
                                    }
                                }
                            }
                        }

                        // 處理臨編日期
                        if (srr.TempDate != "")
                        {
                            DateTime d1;
                            if (DateTime.TryParse(srr.TempDate, out d1))
                            {
                                if (lastUpdateRec == null)
                                {
                                    lastUpdateRec = srr;
                                }
                                else
                                {
                                    DateTime d2;
                                    if (DateTime.TryParse(lastUpdateRec.TempDate, out d2))
                                    {
                                        if (d1 > d2)
                                        {
                                            lastUpdateRec = srr;
                                        }
                                    }
                                }
                            }
                        }

                    }
                }

                // 2024/9/19 客服會議討論後調整：當學生一般狀態，讀取目前班級、年級、座號，學生休學狀態，讀取目前最後一筆異動有日期的年級，當學生延修狀態，最後一筆異動對照的年級，非一般狀態學生不需要填班級、座號。
                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.一般)
                {
                    if (sr.Class != null)
                    {
                        if (sr.Class.GradeYear.HasValue)
                            grade = sr.Class.GradeYear.Value + "";

                        className = sr.Class.Name;

                        if (sr.SeatNo.HasValue)
                            seatNo = sr.SeatNo.Value + "";
                    }
                }


                // 最後一筆
                if (lastUpdateRec != null)
                {
                    classCode = lastUpdateRec.ClassType;

                    // 使用異動的科別名稱反推科別管理找出科別代碼
                    if (deptNameCodeDict.ContainsKey(lastUpdateRec.Department))
                        departmentCode = deptNameCodeDict[lastUpdateRec.Department];

                    // 處理學生年級班級，當異動學年度學期與目前學年度學期相同，讀取目前學生班級座號，不相同讀取學期對照
                    string SY = "", SS = "";
                    if (lastUpdateRec.SchoolYear.HasValue)
                        SY = lastUpdateRec.SchoolYear.Value.ToString();
                    if (lastUpdateRec.Semester.HasValue)
                        SS = lastUpdateRec.Semester.Value.ToString();


                    if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.休學)
                    {
                        grade = lastUpdateRec.GradeYear;
                    }

                    if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.延修)
                    {
                        if (StudSemsHisDict.ContainsKey(sr.ID))
                        {
                            foreach (SemsHistoryInfo sh in StudSemsHisDict[sr.ID])
                            {
                                if (sh.SchoolYear == SY && sh.Semester == SS)
                                {
                                    grade = sh.GradeYear;
                                    break;
                                }
                            }
                        }
                    }
                }

                if (classCode == "")
                {
                    errorReason += "沒有班別資訊，請至學生最新的異動紀錄上新增。";
                }

                if (departmentCode == "")
                {
                    errorReason += "沒有科別資訊，請至學生班級資料上新增，並確認教務作業中代碼設定。";
                }

                //依照班別去對照 上傳類別
                if (updateTypeDict.ContainsKey(classCode))
                {
                    updateType = updateTypeDict[classCode];
                }

                if (updateType == "")
                {
                    errorReason += "沒有上傳類別資訊，請至學生最新的異動紀錄上新增班別資訊，以利系統自動對照產生。";
                }

                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.一般)
                {
                    identityType = "1";

                    if (className == "")
                    {
                        errorReason += "沒有班級名稱，請至學生班級資料或學生學期對照上新增。";
                    }
                }
                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.休學)
                {
                    identityType = "2";

                }
                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.延修)
                {
                    identityType = "3";

                }

                if (grade == "")
                {
                    errorReason += "沒有年級資料，請至學生班級資料或學生學期對照上新增。";
                }

                // 將有不完整資料的學生 加入錯誤清單
                if (classCode == "" || departmentCode == "" || updateType == "" || grade == "" || (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.一般 && className == ""))
                {
                    error_StudentList.Add(sr);

                    errorReasonDict.Add(sr.ID, errorReason);

                    continue;
                }

                //學校代碼
                cs[row_counter, 0].Value = schoolCode;
                //學年度
                cs[row_counter, 1].Value = schoolYear;
                //學期
                cs[row_counter, 2].Value = schoolSemester;
                //名冊別
                cs[row_counter, 3].Value = bookType;
                //班別
                cs[row_counter, 4].Value = classCode;
                //科別學程代碼
                cs[row_counter, 5].Value = departmentCode;
                //上傳類別
                cs[row_counter, 6].Value = updateType;
                //身分證字號
                cs[row_counter, 7].Value = sr.IDNumber;
                //註1
                cs[row_counter, 8].Value = "";
                //身分別
                cs[row_counter, 9].Value = identityType;
                //年級
                cs[row_counter, 10].Value = grade;
                //班級名稱
                cs[row_counter, 11].Value = (_ClassNameMappingDict.ContainsKey(className) ? _ClassNameMappingDict[className] : className);
                //座號
                cs[row_counter, 12].Value = seatNo;
                //實驗班名稱
                cs[row_counter, 13].Value = "";
                //備註
                cs[row_counter, 14].Value = "";

                //複製第一行樣板
                Style s = cs[1, 0].GetStyle();
                if (row_counter > 1)
                {
                    cs[row_counter, 0].SetStyle(s);
                    cs[row_counter, 1].SetStyle(s);
                    cs[row_counter, 2].SetStyle(s);
                    cs[row_counter, 3].SetStyle(s);
                    cs[row_counter, 4].SetStyle(s);
                    cs[row_counter, 5].SetStyle(s);
                    cs[row_counter, 6].SetStyle(s);
                    cs[row_counter, 7].SetStyle(s);
                    cs[row_counter, 8].SetStyle(s);
                    cs[row_counter, 9].SetStyle(s);
                    cs[row_counter, 10].SetStyle(s);
                    cs[row_counter, 11].SetStyle(s);
                    cs[row_counter, 12].SetStyle(s);
                    cs[row_counter, 13].SetStyle(s);
                    cs[row_counter, 14].SetStyle(s);
                }
                row_counter++;
                _BGWClassStudentAbsenceDetail.ReportProgress(50 + (row_counter / target_StudentList.Count) * 40);
            }
            #endregion

            //在「編班名冊」sheet最後一筆的下一列加入End
            for (int i = 0; i < 15; i++)
            {
                cs[row_counter, i].Value = "End";
            }

            // 歸零
            row_counter = 1;


            #region 填寫錯誤報告
            //填寫資料
            foreach (SHSchool.Data.SHStudentRecord sr in error_StudentList)
            {
                identityType = "";
                grade = "";
                className = "";
                seatNo = "";
                departmentCode = "";
                updateType = "";
                errorReason = "";

                // 最後一筆異動
                SHUpdateRecordRecord lastUpdateRec = null;

                // 班別取得規則異動最後一筆有核准日期或臨編日期
                if (updateDict.ContainsKey(sr.ID))
                {
                    // 有核准日期或臨編日期放入
                    foreach (SHUpdateRecordRecord srr in updateDict[sr.ID])
                    {
                        // 處理核准日期
                        if (srr.ADDate != "")
                        {
                            DateTime d1;
                            if (DateTime.TryParse(srr.ADDate, out d1))
                            {
                                if (lastUpdateRec == null)
                                {
                                    lastUpdateRec = srr;
                                }
                                else
                                {
                                    DateTime d2;
                                    if (DateTime.TryParse(lastUpdateRec.ADDate, out d2))
                                    {
                                        if (d1 > d2)
                                        {
                                            lastUpdateRec = srr;
                                        }
                                    }
                                }
                            }
                        }

                        // 處理臨編日期
                        if (srr.TempDate != "")
                        {
                            DateTime d1;
                            if (DateTime.TryParse(srr.TempDate, out d1))
                            {
                                if (lastUpdateRec == null)
                                {
                                    lastUpdateRec = srr;
                                }
                                else
                                {
                                    DateTime d2;
                                    if (DateTime.TryParse(lastUpdateRec.TempDate, out d2))
                                    {
                                        if (d1 > d2)
                                        {
                                            lastUpdateRec = srr;
                                        }
                                    }
                                }
                            }
                        }

                    }
                }

                // 2024/9/19 客服會議討論後調整：當學生一般狀態，讀取目前班級、年級、座號，學生休學狀態，讀取目前最後一筆異動有日期的年級，當學生延修狀態，最後一筆異動對照的年級，非一般狀態學生不需要填班級、座號。
                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.一般)
                {
                    if (sr.Class != null)
                    {
                        if (sr.Class.GradeYear.HasValue)
                            grade = sr.Class.GradeYear.Value + "";

                        className = sr.Class.Name;

                        if (sr.SeatNo.HasValue)
                            seatNo = sr.SeatNo.Value + "";
                    }
                }

                // 最後一筆
                if (lastUpdateRec != null)
                {
                    classCode = lastUpdateRec.ClassType;

                    // 使用異動的科別名稱反推科別管理找出科別代碼
                    if (deptNameCodeDict.ContainsKey(lastUpdateRec.Department))
                        departmentCode = deptNameCodeDict[lastUpdateRec.Department];

                    // 處理學生年級班級，當異動學年度學期與目前學年度學期相同，讀取目前學生班級座號，不相同讀取學期對照
                    string SY = "", SS = "";
                    if (lastUpdateRec.SchoolYear.HasValue)
                        SY = lastUpdateRec.SchoolYear.Value.ToString();
                    if (lastUpdateRec.Semester.HasValue)
                        SS = lastUpdateRec.Semester.Value.ToString();

                    if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.休學)
                    {
                        grade = lastUpdateRec.GradeYear;
                    }

                    if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.延修)
                    {
                        if (StudSemsHisDict.ContainsKey(sr.ID))
                        {
                            foreach (SemsHistoryInfo sh in StudSemsHisDict[sr.ID])
                            {
                                if (sh.SchoolYear == SY && sh.Semester == SS)
                                {
                                    grade = sh.GradeYear;
                                    break;
                                }
                            }
                        }
                    }

                }

                //依照班別去對照 上傳類別
                if (updateTypeDict.ContainsKey(classCode))
                {

                    updateType = updateTypeDict[classCode];
                }

                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.一般)
                {
                    identityType = "1";
                }
                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.休學)
                {
                    identityType = "2";
                }
                if (sr.Status == SHSchool.Data.SHStudentRecord.StudentStatus.延修)
                {
                    identityType = "3";
                }

                //學校代碼
                cs_errorReport[row_counter, 0].Value = schoolCode;
                //學年度
                cs_errorReport[row_counter, 1].Value = schoolYear;
                //學期
                cs_errorReport[row_counter, 2].Value = schoolSemester;
                //名冊別
                cs_errorReport[row_counter, 3].Value = bookType;
                //班別
                cs_errorReport[row_counter, 4].Value = classCode;
                //科別學程代碼
                cs_errorReport[row_counter, 5].Value = departmentCode;
                //上傳類別
                cs_errorReport[row_counter, 6].Value = updateType;
                //身分證字號
                cs_errorReport[row_counter, 7].Value = sr.IDNumber;
                //註1
                cs_errorReport[row_counter, 8].Value = "";
                //身分別
                cs_errorReport[row_counter, 9].Value = identityType;
                //年級
                cs_errorReport[row_counter, 10].Value = grade;
                //班級名稱
                cs_errorReport[row_counter, 11].Value = className;
                //座號
                cs_errorReport[row_counter, 12].Value = seatNo;
                //實驗班名稱
                cs_errorReport[row_counter, 13].Value = "";
                //備註
                cs_errorReport[row_counter, 14].Value = "";
                //錯誤資訊
                cs_errorReport[row_counter, 15].Value = errorReasonDict[sr.ID];

                //複製第一行樣板
                Style s = cs_errorReport[1, 0].GetStyle();
                if (row_counter > 1)
                {
                    cs_errorReport[row_counter, 0].SetStyle(s);
                    cs_errorReport[row_counter, 1].SetStyle(s);
                    cs_errorReport[row_counter, 2].SetStyle(s);
                    cs_errorReport[row_counter, 3].SetStyle(s);
                    cs_errorReport[row_counter, 4].SetStyle(s);
                    cs_errorReport[row_counter, 5].SetStyle(s);
                    cs_errorReport[row_counter, 6].SetStyle(s);
                    cs_errorReport[row_counter, 7].SetStyle(s);
                    cs_errorReport[row_counter, 8].SetStyle(s);
                    cs_errorReport[row_counter, 9].SetStyle(s);
                    cs_errorReport[row_counter, 10].SetStyle(s);
                    cs_errorReport[row_counter, 11].SetStyle(s);
                    cs_errorReport[row_counter, 12].SetStyle(s);
                    cs_errorReport[row_counter, 13].SetStyle(s);
                    cs_errorReport[row_counter, 14].SetStyle(s);
                    cs_errorReport[row_counter, 15].SetStyle(s);
                }
                #endregion

                row_counter++;

                _BGWClassStudentAbsenceDetail.ReportProgress(90 + (row_counter / error_StudentList.Count) * 10);
            }

            //假如沒有錯誤資料的話
            if (error_StudentList.Count == 0)
            {
                _Wk.Worksheets.RemoveAt(1);
            }

        }


        //建立科別代號表
        public void QueryDeptCode()
        {

            QueryHelper _Q = new QueryHelper();
            DataTable dt = _Q.Select("select id,code,name from dept");
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String code = row["code"].ToString();
                if (code == "") code = "NoCode";
                if (!Dept_ref.ContainsKey(id))
                {
                    Dept_ref.Add(id, code);
                }
            }
        }

        //查詢科別代碼
        public String getDeptCode(String id)
        {
            string code = "";

            if (Dept_ref.ContainsKey(id))
            {
                code = Dept_ref[id];

            }
            return code;
        }

    }
}
