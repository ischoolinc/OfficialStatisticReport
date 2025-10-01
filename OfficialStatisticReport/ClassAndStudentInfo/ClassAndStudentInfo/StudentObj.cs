using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ClassAndStudentInfo
{
    class StudentObj
    {
        public string StudentID { get; set; }
        public string StudentName { get; set; }
        public string ClassID { get; set; }
    
        public string Gender { get; set; }
        public string GradeYear { get; set; }
        public string Department { get; set; }
        public string DepartmentID { get; set; }

        public string Status { get; set; }

        public StudentObj(DataRow row)
        {
            StudentID = row["id"].ToString();
            StudentName = row["name"].ToString();
            ClassID = row["ref_class_id"].ToString();      
            Gender = row["gender"].ToString();
            GradeYear = row["grade_year"].ToString();
            Department = row["dept_name"].ToString();
            DepartmentID = row["dept_id"].ToString();
            Status = row["status"].ToString();
        }
    }
}
