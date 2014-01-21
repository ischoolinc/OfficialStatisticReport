using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateRecordReport
{
    class Studentobj
    {
        private string id, name, gender, student_number, grade, dept;

        public Studentobj(string id, string name, string gender, string student_number, string grade, string dept)
        {
            this.id = id;
            this.name = name;
            this.gender = gender;
            this.student_number = student_number;
            this.grade = grade;
            this.dept = dept;
            this.id = id;
            this.codeList = new List<string>();
        }

        //科別
        public string Dept
        {
            get { return dept; }
            set { dept = value; }
        }

        //年級
        public string Grade
        {
            get { return grade; }
            set { grade = value; }
        }

        //學號
        public string Student_number
        {
            get { return student_number; }
            set { student_number = value; }
        }

        //姓名
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        //性別
        public string Gender
        {
            get { return gender; }
            set { gender = value; }
        }

        //系統編號
        public string Id
        {
            get { return id; }
            set { id = value; }
        }

        //異動代碼清單
        private List<string> codeList;

        public List<string> CodeList
        {
            get { return codeList; }
            set { codeList = value; }
        }
    }
}
