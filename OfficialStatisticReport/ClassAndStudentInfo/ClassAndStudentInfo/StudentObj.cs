using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ClassAndStudentInfo
{
    class StudentObj
    {
        public string 學生ID { get; set; }
        public string 學生名字 { get; set; }
        public string 班級ID { get; set; }
        public string 狀態 { get; set; }
        public string 性別 { get; set; }
        public string 年級 { get; set; }
        public string 科別 { get; set; }

        public StudentObj(DataRow row)
        {
            學生ID = row["id"].ToString();
            學生名字 = row["name"].ToString();
            班級ID = row["ref_class_id"].ToString();
            狀態 = row["status"].ToString();
            性別 = row["gender"].ToString();
            年級 = row["grade_year"].ToString();
            科別 = row["dept_name"].ToString();
        }
    }
}
