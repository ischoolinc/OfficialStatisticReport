using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace myTable
{
    class Filter
    {
        private String dept;
        public List<myStudent> error_list, clean_list;
        public Dictionary<String, List<myStudent>> dic_byDept;
        Dictionary<String, String> Dept_ref; //科別代碼對照,key=code,value=name;
        

        public Filter(List<myStudent> list, String dept)
        {
            this.dept = dept;
            Cleaner(list);
            Classify();
            QueryDeptCode();
        }

        //清除總資料異常,正確資料放clean_list,錯誤資料放error_list
        private void Cleaner(List<myStudent> list)
        {
            error_list = new List<myStudent>();
            clean_list = new List<myStudent>();
            foreach (myStudent s in list)
            {
                if (s.Id == "" || s.Name == "" || (s.Gender != "0" && s.Gender != "1") || s.Ref_class_id == "" || s.Class_name == "" || s.Grade_year == "" || s.Dept_name == "")
                {
                    error_list.Add(s);
                }
                else
                {
                    clean_list.Add(s);
                }
            }
        }

        //按科別分類收集,篩選error_list所對應的科別
        private void Classify()
        {
            dic_byDept = new Dictionary<string, List<myStudent>>();
            List<myStudent> new_error_list = new List<myStudent>();

            switch (dept)
            {
                case "職業科":
                    foreach (myStudent s in clean_list)
                    {
                        if (!s.Dept_name.Contains("普通科") && !s.Dept_name.Contains("綜合高中科"))
                        {
                            if (!dic_byDept.ContainsKey(s.Dept_name))
                            {
                                dic_byDept.Add(s.Dept_name, new List<myStudent>());
                            }
                            dic_byDept[s.Dept_name].Add(s);
                        }
                    }

                    foreach (myStudent s in error_list)
                    {
                        if (!s.Dept_name.Contains("普通科") && !s.Dept_name.Contains("綜合高中科"))
                        {
                            new_error_list.Add(s);
                        }
                    }

                    error_list = new_error_list;

                    break;

                default:
                    foreach (myStudent s in clean_list)
                    {
                        if (!dic_byDept.ContainsKey(dept))
                        {
                            dic_byDept.Add(dept, new List<myStudent>());
                        }
                        if (s.Dept_name.Contains(dept))
                        {
                            dic_byDept[dept].Add(s);
                        }
                    }

                    foreach (myStudent s in error_list)
                    {
                        if (s.Dept_name.Contains(dept))
                        {
                            new_error_list.Add(s);
                        }
                    }

                    error_list = new_error_list;

                    break;
            }
        }

        //透過Ref_class_id判斷資料中的班級總數
        public int getClassCount(List<myStudent> list)
        {
            Dictionary<string, List<myStudent>> dic_byClass = new Dictionary<string, List<myStudent>>();
            foreach (myStudent s in list)
            {
                if (!dic_byClass.ContainsKey(s.Ref_class_id))
                {
                    dic_byClass.Add(s.Ref_class_id, new List<myStudent>());
                }
                dic_byClass[s.Ref_class_id].Add(s);
            }
            return dic_byClass.Count;
        }

        //計算傳入的學生物件清單指定的性別數量並回傳
        public int getGenderCount(List<myStudent> list, String gender)
        {
            int count = 0;
            foreach (myStudent s in list)
            {
                if (s.Gender == gender)
                {
                    count++;
                }
            }
            return count;
        }

        //收集符合指定tagId的學生物件並回傳清單(不判斷重複學生)
        public List<myStudent> getListByTagId(List<String> id,List<myStudent> list)
        {
            List<myStudent> collect = new List<myStudent>();
            foreach (myStudent student in list) //從乾淨的總表搜尋
            {
                foreach (String sTag in student.Tag) //每個學生的註記
                {
                    foreach(String sid in id)
                    {
                        if (sid == sTag)
                        {
                            collect.Add(student);
                        }
                    }
                }
            }
            return collect;
        }

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
    }
}





