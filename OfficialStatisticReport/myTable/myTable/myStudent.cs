using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace myTable
{
    class myStudent
    {
        //學生物件存放個資料
        private String id, name, gender, ref_class_id, class_name, grade_year, dept_name;

        
        private List<String> tag;

        public myStudent(String id, String name, String gender, String ref_class_id, String class_name, String grade_year, String dept_name,List<String> tag)
        {
            this.id = id;
            this.name = name;
            this.gender = gender;
            this.ref_class_id = ref_class_id;
            this.class_name = class_name;
            this.grade_year = grade_year;
            this.dept_name = dept_name;
            this.tag = tag;
        }

        //欄位封裝
        public String Id
        {
            get { return id; }
            set { id = value; }
        }

        public String Name
        {
            get { return name; }
            set { name = value; }
        }

        public String Gender
        {
            get { return gender; }
            set { gender = value; }
        }

        public String Ref_class_id
        {
            get { return ref_class_id; }
            set { ref_class_id = value; }
        }

        public String Class_name
        {
            get { return class_name; }
            set { class_name = value; }
        }

        public String Grade_year
        {
            get { return grade_year; }
            set { grade_year = value; }
        }

        public String Dept_name
        {
            get { return dept_name; }
            set { dept_name = value; }
        }

        public List<String> Tag
        {
            get { return tag; }
            set { tag = value; }
        }

       
    }
}
