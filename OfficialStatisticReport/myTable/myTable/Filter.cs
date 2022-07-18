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
        public Dictionary<string, Dictionary<string, List<myStudent>>> deptClassTypeDic;
        public Dictionary<String, List<myStudent>> dept_ClassTypeDic;  //科別_班別
        Dictionary<String, String> Dept_ref; //科別代碼對照,key=code,value=name;
        public Dictionary<string, string> ClassTypeCodeDic;


        public Filter(List<myStudent> list, String dept)
        {
            this.dept = dept;
            LoadClassTypeCodeDic();
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
                //2022-07-13 Cynthia 班別不在對照內，放到錯誤清單
                if (s.Id == "" || s.Name == "" || (s.Gender != "0" && s.Gender != "1") || s.Ref_class_id == "" || s.Class_name == "" || s.Grade_year == "" || s.Dept_name == "" || !ClassTypeCodeDic.ContainsKey(s.Class_Type))
                {
                    error_list.Add(s);
                }
                else
                {
                    clean_list.Add(s);
                }
            }
        }

        /// <summary>
        /// 載入班別代碼對照
        /// </summary>
        private void LoadClassTypeCodeDic()
        {
            ClassTypeCodeDic = new Dictionary<string, string>();
            ClassTypeCodeDic.Add("1", "");  //日間部 //客服表示日間部不用填入
            ClassTypeCodeDic.Add("2", "夜間部");
            ClassTypeCodeDic.Add("3", "實用技能學程(一般班)");
            ClassTypeCodeDic.Add("4", "建教班");
            ClassTypeCodeDic.Add("6", "產學訓合作計畫班(產學合作班)");
            ClassTypeCodeDic.Add("7", "重點產業班/台德菁英班/雙軌旗艦訓練計畫專班");
            ClassTypeCodeDic.Add("8", "建教僑生專班");
            ClassTypeCodeDic.Add("9", "實驗班");
            ClassTypeCodeDic.Add("01", "進修部(核定班)");
            ClassTypeCodeDic.Add("02", "編制班");
            ClassTypeCodeDic.Add("03", "自給自足班");
            ClassTypeCodeDic.Add("04", "員工進修班");
            ClassTypeCodeDic.Add("05", "重點產業班");
            ClassTypeCodeDic.Add("06", "產業人力套案專班");
        }
        //按科別分類收集,篩選error_list所對應的科別
        private void Classify()
        {
            dic_byDept = new Dictionary<string, List<myStudent>>();
            deptClassTypeDic = new Dictionary<string, Dictionary<string, List<myStudent>>>();
            dept_ClassTypeDic = new Dictionary<string, List<myStudent>>();
            List<myStudent> new_error_list = new List<myStudent>();

            switch (dept)
            {
                case "職業科":
                    foreach (myStudent s in clean_list)
                    {
                        //科別⊕班別
                        string key = s.Dept_name + "⊕" + s.Class_Type;

                        if (!s.Dept_name.Contains("普通科") && !s.Dept_name.Contains("綜合高中科"))
                        {
                            if (!dic_byDept.ContainsKey(s.Dept_name))
                            {
                                dic_byDept.Add(s.Dept_name, new List<myStudent>());
                            }
                            dic_byDept[s.Dept_name].Add(s);

                            if (!dept_ClassTypeDic.ContainsKey(key))
                            {
                                dept_ClassTypeDic.Add(key, new List<myStudent>());
                            }
                            dept_ClassTypeDic[key].Add(s);
                            #region deptClassTypeDic
                            if (!deptClassTypeDic.ContainsKey(s.Dept_name))
                            {
                                deptClassTypeDic.Add(s.Dept_name, new Dictionary<string, List<myStudent>>());
                            }
                            if (!deptClassTypeDic[s.Dept_name].ContainsKey(s.Class_Type))
                                deptClassTypeDic[s.Dept_name].Add(s.Class_Type, new List<myStudent>());
                            deptClassTypeDic[s.Dept_name][s.Class_Type].Add(s);
                            #endregion

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
                        //科別⊕班別
                        string key = s.Dept_name + "⊕" + s.Class_Type;

                        if (!dic_byDept.ContainsKey(dept))
                        {
                            dic_byDept.Add(dept, new List<myStudent>());
                        }
                        if (!dept_ClassTypeDic.ContainsKey(key) && s.Dept_name.Contains(dept))
                        {
                            dept_ClassTypeDic.Add(key, new List<myStudent>());
                        }

                        if (s.Dept_name.Contains(dept))
                        {
                            dept_ClassTypeDic[key].Add(s);
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
        public List<myStudent> getListByTagId(List<String> id, List<myStudent> list)
        {
            List<myStudent> collect = new List<myStudent>();
            foreach (myStudent student in list) //從乾淨的總表搜尋
            {
                foreach (String sTag in student.Tag) //每個學生的註記
                {
                    foreach (String sid in id)
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





