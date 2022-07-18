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
        private String id, name, gender, ref_class_id, class_name, grade_year, dept_name,county,before_school_location, class_type;

        
        private List<String> tag;

        public myStudent(String id, String name, String gender, String ref_class_id, String class_name, String grade_year, String dept_name, String county,String before_school_location, String class_type, List<String> tag)
        {
            this.id = id;
            this.name = name;
            this.gender = gender;
            this.ref_class_id = ref_class_id;
            this.class_name = class_name;
            this.grade_year = grade_year;
            this.dept_name = dept_name;
            this.county = county;
            this.before_school_location = before_school_location;
            this.class_type = class_type;
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

        public String County
        {
            get { return county; }
            set { county = value; }
        }

        public String Before_School_Location
        {
            get { return before_school_location; }
            set { before_school_location = value; }
        }

        public String Class_Type
        {
            get { return class_type; }
            set { class_type = value; }
        }

        public List<String> Tag
        {
            get { return tag; }
            set { tag = value; }
        }

       
    }
}
