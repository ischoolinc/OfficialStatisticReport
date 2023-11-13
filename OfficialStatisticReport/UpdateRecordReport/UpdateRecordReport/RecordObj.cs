using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateRecordReport
{
    class RecordObj
    {
        private string id, student_id, name, student_number, gender, grade, dept, code,status;
        private bool delay;

        public RecordObj(DataRow row)
        {
            this.id = row["id"].ToString();
            this.student_id = row["ref_student_id"].ToString();
            this.name = row["ss_name"].ToString();
            this.student_number = row["ss_student_number"].ToString();
            this.gender = row["ss_gender"].ToString();
            this.grade = row["ss_grade_year"].ToString();
            this.dept = row["ss_dept"].ToString();
            this.code = row["update_code"].ToString();
            this.status = row["status"].ToString();
            this.delay = false;
        }

        public string Id
        {
            get { return id; }
            set { id = value; }
        }

        public string Student_id
        {
            get { return student_id; }
            set { student_id = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Student_number
        {
            get { return student_number; }
            set { student_number = value; }
        }

        public string Gender
        {
            get { return gender; }
            set { gender = value; }
        }

        public string Grade
        {
            get { return grade; }
            set { grade = value; }
        }

        public string Dept
        {
            get { return dept; }
            set { dept = value; }
        }

        public string Code
        {
            get { return code; }
            set { code = value; }
        }
        public string Status
        {
            get { return status; }
            set { status = value; }
        }
        public bool Delay
        {
            get { return delay; }
            set { delay = value; }
        }
    }
}
