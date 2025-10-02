using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;


namespace ClassAndStudentInfo
{
    class GraduateStudentObj
    {
        public string StudentID { get; set; }
        public string StudentName { get; set; }
        public string Gender { get; set; }
        public string Department { get; set; }
        public string DepartmentID { get; set; }

        public GraduateStudentObj(DataRow row)
        {
            StudentID = row["ref_student_id"].ToString();
            StudentName = row["ss_name"].ToString();
            Gender = row["ss_gender"].ToString();
            Department = row["ss_dept"].ToString();
            DepartmentID = row["dept_id"].ToString();
        }
    }
}
