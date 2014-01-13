using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;


namespace ClassAndStudentInfo
{
    class GraduateStudentObj
    {
        public string 學生ID { get; set; }
        public string 學生名字 { get; set; }
        public string 性別 { get; set; }
        public string 科別 { get; set; }

        public GraduateStudentObj(DataRow row)
        {
            學生ID = row["ref_student_id"].ToString();
            學生名字 = row["ss_name"].ToString();
            性別 = row["ss_gender"].ToString();
            科別 = row["ss_dept"].ToString();
        }
    }
}
